"""Shared models, training, prediction and metric utilities for the IDS polymorphic-robustness project.
Imported by notebooks 02, 04 and 05 so that every experiment uses identical settings."""
import os, json, time, random
import numpy as np
import torch
import torch.nn as nn
import joblib
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (accuracy_score, precision_recall_fscore_support, confusion_matrix)

MODEL_NAMES = ['RandomForest', 'MLP', 'CNN1D', 'LSTM']
DL_MODELS = ['MLP', 'CNN1D', 'LSTM']

RF_HPARAMS = {'n_estimators': 150, 'n_jobs': -1}
DL_HPARAMS = {'lr': 1e-3, 'batch_size': 512, 'max_epochs': 40, 'patience': 5, 'min_delta': 1e-4,
              'eval_batch_size': 4096}


# ------------------------------------------------------------------ reproducibility
def set_seed(seed):
    random.seed(seed); np.random.seed(seed)
    torch.manual_seed(seed); torch.cuda.manual_seed_all(seed)
    torch.backends.cudnn.deterministic = True
    torch.backends.cudnn.benchmark = False
    os.environ['PYTHONHASHSEED'] = str(seed)


def get_device():
    return torch.device('cuda' if torch.cuda.is_available() else 'cpu')


# ------------------------------------------------------------------ models
class MLP(nn.Module):
    def __init__(self, n_features, n_classes, dropout=0.3):
        super().__init__()
        self.net = nn.Sequential(
            nn.Linear(n_features, 128), nn.ReLU(), nn.Dropout(dropout),
            nn.Linear(128, 64), nn.ReLU(), nn.Dropout(dropout),
            nn.Linear(64, 32), nn.ReLU(), nn.Dropout(dropout),
            nn.Linear(32, n_classes))

    def forward(self, x):
        return self.net(x)


class CNN1D(nn.Module):
    def __init__(self, n_features, n_classes, dropout=0.3):
        super().__init__()
        self.features = nn.Sequential(
            nn.Conv1d(1, 32, kernel_size=3, padding=1), nn.ReLU(), nn.MaxPool1d(2),
            nn.Conv1d(32, 64, kernel_size=3, padding=1), nn.ReLU(), nn.MaxPool1d(2))
        flat = 64 * (n_features // 4)
        self.head = nn.Sequential(nn.Flatten(), nn.Linear(flat, 64), nn.ReLU(), nn.Dropout(dropout),
                                  nn.Linear(64, n_classes))

    def forward(self, x):
        return self.head(self.features(x.unsqueeze(1)))       # (B, F) -> (B, 1, F)


class LSTMNet(nn.Module):
    def __init__(self, n_features, n_classes, hidden=64, dropout=0.2):
        super().__init__()
        self.lstm = nn.LSTM(input_size=1, hidden_size=hidden, num_layers=1, batch_first=True)
        self.head = nn.Sequential(nn.Dropout(dropout), nn.Linear(hidden, n_classes))

    def forward(self, x):
        _, (h, _) = self.lstm(x.unsqueeze(-1))                # (B, F) -> (B, F, 1); h: (1, B, H)
        return self.head(h[-1])


def build_dl_model(name, n_features, n_classes):
    if name == 'MLP':
        return MLP(n_features, n_classes)
    if name == 'CNN1D':
        return CNN1D(n_features, n_classes)
    if name == 'LSTM':
        return LSTMNet(n_features, n_classes)
    raise ValueError(f'unknown DL model {name}')


# ------------------------------------------------------------------ training
def train_rf(X, y, seed):
    rf = RandomForestClassifier(random_state=seed, **RF_HPARAMS)
    t0 = time.time()
    rf.fit(X, y)
    return rf, {'train_time_s': time.time() - t0}


@torch.no_grad()
def _eval_loss(model, X, y, criterion, bs):
    model.eval()
    total, n = 0.0, 0
    for i in range(0, len(X), bs):
        xb, yb = X[i:i + bs], y[i:i + bs]
        total += criterion(model(xb), yb).item() * len(xb); n += len(xb)
    return total / n


def train_dl(name, X_tr, y_tr, X_va, y_va, seed, n_classes, device=None, class_weights=None, verbose=True):
    """Adam + cross-entropy, early stopping on validation loss with best-weight restore."""
    device = device or get_device()
    set_seed(seed)
    hp = DL_HPARAMS
    Xtr = torch.as_tensor(X_tr, dtype=torch.float32, device=device)
    ytr = torch.as_tensor(y_tr, dtype=torch.long, device=device)
    Xva = torch.as_tensor(X_va, dtype=torch.float32, device=device)
    yva = torch.as_tensor(y_va, dtype=torch.long, device=device)

    model = build_dl_model(name, Xtr.shape[1], n_classes).to(device)
    w = None
    if class_weights is not None:
        w = torch.as_tensor([class_weights[i] for i in range(n_classes)], dtype=torch.float32, device=device)
    criterion = nn.CrossEntropyLoss(weight=w)
    eval_criterion = nn.CrossEntropyLoss(weight=w)
    opt = torch.optim.Adam(model.parameters(), lr=hp['lr'])
    gen = torch.Generator(device='cpu').manual_seed(seed)

    best_loss, best_state, best_epoch, wait = float('inf'), None, 0, 0
    history = {'train_loss': [], 'val_loss': [], 'val_acc': []}
    t0 = time.time()
    for epoch in range(1, hp['max_epochs'] + 1):
        model.train()
        perm = torch.randperm(len(Xtr), generator=gen).to(device)
        run, n = 0.0, 0
        for i in range(0, len(Xtr), hp['batch_size']):
            idx = perm[i:i + hp['batch_size']]
            xb, yb = Xtr[idx], ytr[idx]
            opt.zero_grad(set_to_none=True)
            loss = criterion(model(xb), yb)
            loss.backward()
            opt.step()
            run += loss.item() * len(idx); n += len(idx)
        tr_loss = run / n
        va_loss = _eval_loss(model, Xva, yva, eval_criterion, hp['eval_batch_size'])
        va_pred = predict_dl(model, Xva, device=device)
        va_acc = float((va_pred == y_va).mean())
        history['train_loss'].append(tr_loss); history['val_loss'].append(va_loss); history['val_acc'].append(va_acc)
        improved = va_loss < best_loss - hp['min_delta']
        if improved:
            best_loss, best_epoch, wait = va_loss, epoch, 0
            best_state = {k: v.detach().cpu().clone() for k, v in model.state_dict().items()}
        else:
            wait += 1
        if verbose:
            print(f'    [{name} seed {seed}] epoch {epoch:02d}  train {tr_loss:.4f}  val {va_loss:.4f}  '
                  f'val_acc {va_acc:.4f}{"  *" if improved else ""}')
        if wait >= hp['patience']:
            break
    model.load_state_dict(best_state)
    info = {'train_time_s': time.time() - t0, 'best_epoch': best_epoch, 'epochs_run': epoch,
            'best_val_loss': best_loss, 'history': history}
    del Xtr, ytr, Xva, yva
    if device.type == 'cuda':
        torch.cuda.empty_cache()
    return model, info


@torch.no_grad()
def predict_proba_dl(model, X, device=None, bs=None):
    device = device or get_device()
    bs = bs or DL_HPARAMS['eval_batch_size']
    model.eval()
    out = []
    for i in range(0, len(X), bs):
        xb = X[i:i + bs]
        if not torch.is_tensor(xb):
            xb = torch.as_tensor(np.asarray(xb), dtype=torch.float32)
        out.append(torch.softmax(model(xb.to(device)), dim=1).cpu().numpy())
    return np.concatenate(out, axis=0)


def predict_dl(model, X, device=None, bs=None):
    return predict_proba_dl(model, X, device, bs).argmax(axis=1)


def predict(model_name, model, X, device=None):
    if model_name == 'RandomForest':
        return model.predict(X)
    return predict_dl(model, X, device=device)


# ------------------------------------------------------------------ persistence
def model_path(model_dir, model_name, seed):
    ext = 'joblib' if model_name == 'RandomForest' else 'pt'
    return os.path.join(model_dir, f'{model_name}_seed{seed}.{ext}')


def save_model(model_name, model, path, n_features=None, n_classes=None):
    if model_name == 'RandomForest':
        joblib.dump(model, path, compress=3)
    else:
        torch.save({'state_dict': model.state_dict(), 'model_name': model_name,
                    'n_features': n_features, 'n_classes': n_classes}, path)


def load_model(model_name, path, device=None):
    if model_name == 'RandomForest':
        return joblib.load(path)
    device = device or get_device()
    ckpt = torch.load(path, map_location=device)
    m = build_dl_model(model_name, ckpt['n_features'], ckpt['n_classes']).to(device)
    m.load_state_dict(ckpt['state_dict'])
    m.eval()
    return m


# ------------------------------------------------------------------ metrics
def compute_metrics(y_true, y_pred, class_names, benign_id=0):
    """Multi-class metrics + binary (attack vs benign) FPR / detection rate."""
    y_true = np.asarray(y_true); y_pred = np.asarray(y_pred)
    labels = list(range(len(class_names)))
    p_c, r_c, f_c, s_c = precision_recall_fscore_support(y_true, y_pred, labels=labels, zero_division=0)
    p_m, r_m, f_m, _ = precision_recall_fscore_support(y_true, y_pred, labels=labels, average='macro', zero_division=0)
    p_w, r_w, f_w, _ = precision_recall_fscore_support(y_true, y_pred, labels=labels, average='weighted', zero_division=0)
    cm = confusion_matrix(y_true, y_pred, labels=labels)

    true_att = y_true != benign_id
    pred_att = y_pred != benign_id
    tn = int((~true_att & ~pred_att).sum()); fp = int((~true_att & pred_att).sum())
    fn = int((true_att & ~pred_att).sum()); tp = int((true_att & pred_att).sum())
    m = {
        'accuracy': float(accuracy_score(y_true, y_pred)),
        'precision_macro': float(p_m), 'recall_macro': float(r_m), 'f1_macro': float(f_m),
        'precision_weighted': float(p_w), 'recall_weighted': float(r_w), 'f1_weighted': float(f_w),
        'fpr': fp / (fp + tn) if (fp + tn) else float('nan'),               # benign flagged as any attack
        'detection_rate': tp / (tp + fn) if (tp + fn) else float('nan'),    # binary attack recall
        'evasion_rate': fn / (tp + fn) if (tp + fn) else float('nan'),      # attacks predicted Benign
        'n_samples': int(len(y_true)),
    }
    for i, c in enumerate(class_names):
        m[f'precision_{c}'] = float(p_c[i]); m[f'recall_{c}'] = float(r_c[i])
        m[f'f1_{c}'] = float(f_c[i]); m[f'support_{c}'] = int(s_c[i])
    return m, cm


def attack_set_metrics(y_true, y_pred, class_names, benign_id=0):
    """For sets that contain ONLY attack flows (LLM / scripted variants):
    per-class recall (correct attack class), detection (any attack label) and evasion (labelled Benign)."""
    y_true = np.asarray(y_true); y_pred = np.asarray(y_pred)
    out = {'n_samples': int(len(y_true)),
           'recall_overall': float((y_true == y_pred).mean()) if len(y_true) else float('nan'),
           'detection_rate': float((y_pred != benign_id).mean()) if len(y_true) else float('nan'),
           'evasion_rate': float((y_pred == benign_id).mean()) if len(y_true) else float('nan')}
    for i, c in enumerate(class_names):
        if i == benign_id:
            continue
        mask = y_true == i
        n = int(mask.sum())
        out[f'n_{c}'] = n
        out[f'recall_{c}'] = float((y_pred[mask] == i).mean()) if n else float('nan')
        out[f'detection_{c}'] = float((y_pred[mask] != benign_id).mean()) if n else float('nan')
        out[f'evasion_{c}'] = float((y_pred[mask] == benign_id).mean()) if n else float('nan')
    return out


def mean_std_table(df, group_cols, metric_cols, decimals=4):
    """Return (numeric mean/std frame, pretty 'mean ± std' string frame)."""
    g = df.groupby(group_cols)[metric_cols]
    mean, std = g.mean(), g.std(ddof=1).fillna(0.0)
    pretty = mean.copy().astype(object)
    for c in metric_cols:
        pretty[c] = [f'{a:.{decimals}f} ± {b:.{decimals}f}' for a, b in zip(mean[c], std[c])]
    num = mean.add_suffix('_mean').join(std.add_suffix('_std'))
    return num, pretty
