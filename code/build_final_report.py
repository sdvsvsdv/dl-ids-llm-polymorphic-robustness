"""Build Final_Project_Report_IDS_Polymorphic_Robustness.docx from the project's result artifacts.

Run from the repository root:  py -3.13 code/build_final_report.py
"""
import csv
import json
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
RES = ROOT / "results"
FIG = RES / "figures"
OUT = ROOT / "Final_Project_Report_IDS_Polymorphic_Robustness.docx"

prep = json.loads((RES / "01_preprocessing_report.json").read_text())
feat = json.loads((ROOT / "config" / "feature_config.json").read_text())
proj = json.loads((ROOT / "config" / "project_config.json").read_text())
gen = json.loads((ROOT / "config" / "variant_generation_config.json").read_text())
val = {r["source"]: r for r in csv.DictReader(open(RES / "03_validation_summary.csv", encoding="utf-8"))}
summary = list(csv.DictReader(open(RES / "FINAL_summary_table.csv", encoding="utf-8")))
S = {r["Model"]: r for r in summary}

CLASSES = proj["class_names"]
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
HEADER_FILL = "1F3A5F"
ALT_FILL = "EEF2F7"


def m(model, col):
    """Mean part of a 'mean ± std' cell."""
    return float(S[model][col].split("±")[0])


# ----------------------------------------------------------------------------- helpers
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.PORTRAIT
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Inches(1))

st = doc.styles["Normal"]
st.font.name = "Calibri"
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.15
for name, size in (("Heading 1", 18), ("Heading 2", 14), ("Heading 3", 12)):
    h = doc.styles[name]
    h.font.name = "Calibri"
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        h.element.rPr.rFonts.attrib.pop(qn(a), None)  # let explicit font win
    h.font.size = Pt(size)
    h.font.color.rgb = NAVY
    h.font.bold = True
    h.paragraph_format.space_before = Pt(14 if name != "Heading 3" else 10)
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.keep_with_next = True

fig_no = [0]
tab_no = [0]


def para(text="", bold_lead=None, italic=False, align=None, size=None):
    p = doc.add_paragraph()
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
    r = p.add_run(text)
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if align is not None:
        p.alignment = align
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def bullets(items, style="List Bullet"):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            p.add_run(it[0]).bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def caption(text, kind):
    counter = fig_no if kind == "Figure" else tab_no
    counter[0] += 1
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{kind} {counter[0]}. ")
    r.bold = True
    r.font.size = Pt(9.5)
    r2 = p.add_run(text)
    r2.italic = True
    r2.font.size = Pt(9.5)
    if kind == "Table":
        p.paragraph_format.keep_with_next = True
    return p


def table(header, rows, widths=None, font=9.5, caption_text=None, bold_first_col=True):
    if caption_text:
        caption(caption_text, "Table")
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.bold = True
        r.font.size = Pt(font)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c, HEADER_FILL)
    repeat_header(t.rows[0])
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v))
            r.font.size = Pt(font)
            r.bold = bold_first_col and i == 0
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 0 else WD_ALIGN_PARAGRAPH.CENTER
            if ri % 2 == 1:
                shade(cells[i], ALT_FILL)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.line_spacing = 1.0
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def figure(path, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(str(path), width=Inches(6.5))
    caption(text, "Figure")


def fmt(n):
    return f"{n:,}"


# ----------------------------------------------------------------------------- title page
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Egypt-Japan University of Science and Technology (E-JUST)")
r.bold = True
r.font.size = Pt(12)
para("Data Security Course  |  Advisory Project No. 5  |  Final Project Report", align=WD_ALIGN_PARAGRAPH.CENTER, size=11)

h = doc.add_heading("Evaluating the Robustness and Resilience of Deep Learning Intrusion Detection Systems "
                    "Against LLM-Generated Polymorphic Attack Traffic", level=1)
h.alignment = WD_ALIGN_PARAGRAPH.CENTER

table(["Team member", "Student ID"],
      [("George Ehab", "320230042"), ("Ahmed Abdelkader", "320230035"), ("Anas Hany", "320230040"),
       ("Omar Momtaz", "320230071"), ("Eyad Mostafa", "320230028")],
      widths=[3.0, 2.0], font=10.5, bold_first_col=False)

# ----------------------------------------------------------------------------- 1 Abstract
n_llm_in, n_llm_kept, n_llm_drop = int(val["llm"]["input_rows"]), int(val["llm"]["kept_after_validation"]), int(val["llm"]["discarded"])
doc.add_heading("1. Abstract", level=2)
para("Deep learning (DL) intrusion detection systems (IDS) report near-perfect accuracy on public benchmarks, but those "
     "scores are measured on attack patterns the model has already seen in similar form. Large language models (LLMs) make "
     "it cheap for an attacker to produce many slightly different versions of the same attack, a property known as "
     "polymorphism. This project measured that risk in a controlled, reproducible setting.")
para(f"We trained three DL models (MLP, 1D-CNN, LSTM) and a Random Forest baseline on {fmt(prep['cleaning']['final_rows'])} "
     "cleaned and de-duplicated CIC-IDS2017 flows covering Benign traffic, DoS Hulk, PortScan and FTP/SSH brute force. "
     f"We used an LLM to rewrite attack flows at the CICFlowMeter feature level under four behavioural profiles; after a "
     f"five-step automated validation, {n_llm_kept} of {n_llm_in} candidate rows were kept. Every experiment was repeated "
     "over three seeds (42, 102, 2025).")
para("All four models scored highly on clean data (F1-macro 0.966-0.999). On held-out LLM variants, the Random Forest lost "
     "0.207 attack recall (22.5 % evasion) and the LSTM lost 0.164 (15.7 % evasion), while the MLP (0.025, 2.2 %) and 1D-CNN "
     "(0.035, 2.8 %) remained resilient. LLM variants evaded every model more often than class-matched scripted noise. "
     "Adversarial retraining with only 20 % of the variant pool restored Random Forest recall from 0.792 to 0.982, recovering "
     "91.8 % of the lost recall with no measurable cost on clean traffic.")

# ----------------------------------------------------------------------------- 2 Introduction
doc.add_heading("2. Introduction and Motivation", level=2)
doc.add_heading("2.1 Why this topic matters", level=3)
para("Network intrusion detection systems are a core defensive layer. Modern research systems use neural networks that learn "
     "from flow-level statistics such as flow duration, packet counts, packet sizes and inter-arrival times, and on benchmark "
     "datasets they routinely report accuracy above 99 %. Two factors make these numbers less reassuring than they look:")
bullets([("Benchmarks are static. ", "The test set comes from the same capture environment and the same attack tools as the "
          "training set, whereas real attackers change tools, timing and parameters."),
         ("Generative AI lowers the cost of variation. ", "An attacker no longer needs deep expertise to rewrite an attack, "
          "change its pacing or randomise its parameters; an LLM can propose many such changes in seconds.")])
para("Together these motivated the question we investigated: how well do DL IDS models hold up when attack traffic is "
     "rewritten by an LLM?")
doc.add_heading("2.2 What polymorphic attack traffic means", level=3)
para("In malware, a polymorphic sample changes its surface form while keeping its behaviour. We applied the same idea to "
     "network traffic. A polymorphic variant of a port scan is still a port scan, but it may probe more slowly, with "
     "different packet sizes or different pacing. A flow-based IDS sees these changes as shifts in feature values.")
doc.add_heading("2.3 What we did", level=3)
para("We trained IDS models on original data, generated and validated LLM-based variants of three attack families, tested "
     "the models on held-out variants, measured the recall drop and evasion rate, compared the LLM against a scripted-noise "
     "control, and finally tested whether adversarial retraining with 20 % or 50 % of the variant pool restored detection.")

# ----------------------------------------------------------------------------- 3 Problem
doc.add_heading("3. Problem Statement", level=2)
para("DL-based IDS models are usually evaluated only on data that resembles their training data. Before this study it was "
     "unclear how much detection performance degrades when attack traffic is systematically varied by an LLM, and whether a "
     "lightweight retraining strategy can restore it.")

# ----------------------------------------------------------------------------- 4 Objectives
doc.add_heading("4. Objectives, Research Questions and Hypotheses", level=2)
doc.add_heading("4.1 Objectives (all achieved)", level=3)
bullets(["O1. We built and evaluated baseline IDS models (MLP, 1D-CNN, LSTM and a Random Forest baseline) on CIC-IDS2017.",
         "O2. We generated and validated polymorphic variants of DoS Hulk, PortScan and BruteForce flows using an LLM.",
         "O3. We quantified the robustness of each model on held-out variants, against a scripted-noise control.",
         "O4. We tested adversarial retraining at two augmentation levels and measured its cost on clean traffic."])
doc.add_heading("4.2 Research questions", level=3)
table(["ID", "Question"], [
    ("RQ1", "How much does detection performance (recall, F1) drop when models face LLM-generated variants compared with the original test set?"),
    ("RQ2", "What share of the variants is classified as benign (the evasion rate), and does it differ across attack types?"),
    ("RQ3", "Do the DL models differ from the Random Forest baseline, and from each other, in robustness?"),
    ("RQ4", "Does adding a portion of the variants to the training data recover detection without harming results on the original data?")],
    widths=[0.6, 5.9], caption_text="Research questions addressed by the study.")
doc.add_heading("4.3 Hypotheses and outcome", level=3)
bullets([("H1 — supported. ", "Every model showed a measurable recall drop on the variants (0.025 to 0.207)."),
         ("H2 — supported. ", "The drop depended on the attack type (mean evasion 18.3 % for PortScan vs. 6.4 % for BruteForce) "
          "and on the model family."),
         ("H3 — largely supported. ", "Retraining recovered most of the lost recall for the Random Forest (91.8 %) with no clean-data "
          "cost; for the LSTM the recovery was partial and required 50 % augmentation.")])

# ----------------------------------------------------------------------------- 5 Background
doc.add_heading("5. Background", level=2)
doc.add_heading("5.1 Flow-based intrusion detection", level=3)
para("A flow is a group of packets sharing the same source and destination addresses, ports and protocol within a time "
     "window. CICFlowMeter converts raw packet captures into flow records with roughly 80 statistical features, and a "
     "classifier labels each flow as benign or as one of several attack classes. This is the setting of CIC-IDS2017.")
doc.add_heading("5.2 Models for IDS", level=3)
bullets([("MLP: ", "a feed-forward network applied directly to the feature vector; simple, fast and a common starting point."),
         ("1D-CNN: ", "treats the feature vector as a short sequence and learns local patterns with convolution filters."),
         ("LSTM: ", "a recurrent network (Hochreiter and Schmidhuber, 1997) that processes the features as a sequence."),
         ("Random Forest: ", "a tree ensemble (Breiman, 2001), strong on tabular data, used as the non-DL reference.")])
doc.add_heading("5.3 Adversarial robustness and adversarial training", level=3)
para("Small, deliberate input changes can cause confident misclassification, and the standard defense is adversarial training, "
     "in which perturbed examples are added to the training set (Goodfellow et al., 2015). We applied this idea with "
     "perturbations produced by an LLM rewriting attack flows rather than by a gradient-based attack, which is closer to a "
     "practical attacker who has a chat interface but no access to the model's internals.")
doc.add_heading("5.4 LLMs as generators of variation", level=3)
para("LLMs can read rows of numeric features and produce plausible new rows that follow instructions such as \"keep the "
     "behaviour, but make it slower\". Their output is not guaranteed to be valid, so every generated row passed the automated "
     "validation pipeline described in Section 8.4.")

# ----------------------------------------------------------------------------- 6 Scope
doc.add_heading("6. Scope and Assumptions", level=2)
para("In scope: ", bold_lead=None)
bullets(["One public dataset, CIC-IDS2017 (MachineLearningCSV release; Tuesday, Wednesday and Friday-afternoon PortScan files).",
         "Three attack families: DoS Hulk, PortScan and BruteForce (FTP-Patator + SSH-Patator merged).",
         "Three DL models (MLP, 1D-CNN, LSTM) and one Random Forest baseline.",
         "Polymorphic variants generated at the flow-feature level; one defense (adversarial retraining)."])
para("Out of scope: generating real packets or running real attacks, deploying a real-time IDS, gradient-based attacks, and "
     "large hyperparameter searches. We assumed that a feature-level variant is an acceptable proxy for a changed attack "
     "provided its values stay inside realistic, internally consistent ranges.")

# ----------------------------------------------------------------------------- 7 Related work
doc.add_heading("7. Related Work", level=2)
para("CIC-IDS2017 (Sharafaldin et al., 2018) provides labelled flows covering several modern attacks; UNSW-NB15 (Moustafa and "
     "Slay, 2015) is a widely used alternative. Both have known quality issues such as labelling errors and duplicates "
     "(Engelen et al., 2021), which we addressed during preprocessing.", bold_lead="Benchmark datasets. ")
para("Many studies apply MLP, CNN and LSTM models to these datasets and report very high scores; random splits often place "
     "near-duplicates of training flows in the test set, overstating real-world performance.", bold_lead="DL for intrusion detection. ")
para("Goodfellow et al. (2015) showed that models can be fooled by small perturbations and that training on such examples "
     "improves robustness. Network security adds the constraint that perturbed flows must stay valid, which motivated our "
     "range clipping and consistency checks.", bold_lead="Adversarial machine learning. ")
para("LLMs are studied both as tools for defenders and as tools for attackers. Our project examined the defender's side of "
     "that capability: whether current IDS models cope with it.", bold_lead="LLMs in security. ")

# ----------------------------------------------------------------------------- 8 Methodology
doc.add_heading("8. Methodology", level=2)
para("We implemented the study as a five-stage, fully reproducible pipeline. Each stage reads the artifacts of the previous "
     "one from a shared project directory and writes its own configuration and result files.")

doc.add_heading("8.1 Stage 1: Data preparation", level=3)
c = prep["cleaning"]
para(f"We loaded {fmt(c['initial_rows'])} flows from the three selected CIC-IDS2017 files after mapping labels to four classes "
     "(other Wednesday DoS variants and Heartbleed were excluded). Cleaning then removed "
     f"{fmt(c['removed_nonfinite'])} rows with infinite or missing values, {fmt(c['removed_negative'])} rows with physically "
     f"impossible negative values (mostly in Flow IAT Min), {fmt(c['removed_exact_duplicates'])} exact duplicates and "
     f"{fmt(c['removed_label_conflicts'])} rows whose identical features carried conflicting labels. Duplicates were removed "
     "before splitting to prevent train/test leakage.")
table(["Class", "Before cleaning", "After cleaning & de-duplication", "Duplicates removed", "Share of clean data"],
      [(k, fmt(c["initial_counts"][k]), fmt(c["final_counts"][k]), fmt(c["exact_duplicates_by_class"][k]),
        f"{prep['splits']['realistic_mix_percent_full_clean_data'][k]:.2f} %") for k in CLASSES] +
      [("Total", fmt(c["initial_rows"]), fmt(c["final_rows"]), fmt(c["removed_exact_duplicates"]), "100.00 %")],
      widths=[1.2, 1.2, 1.7, 1.2, 1.2], caption_text="Dataset size per class before and after cleaning (01_preprocessing_report.json).")

sp = prep["splits"]
para("We then applied a stratified 70/15/15 train/validation/test split (data seed 42). To control the heavy class imbalance, "
     f"the training set was rebalanced by capping each attack class at {fmt(proj['balancing']['attack_cap_train'])} rows and "
     "down-sampling Benign to a 1:1 benign-to-attack ratio; the validation set was balanced in the same way. The test set "
     "kept the realistic class mix so that FPR and recall reflect operational conditions. Balanced class weights were also "
     "passed to the loss function.")
rows = []
for k in CLASSES:
    rows.append((k, fmt(sp["before_balancing"]["train"][k]), fmt(sp["after_balancing"]["train"][k]),
                 fmt(sp["before_balancing"]["val"][k]), fmt(sp["after_balancing"]["val"][k]), fmt(sp["after_balancing"]["test"][k])))
tot = lambda s, p: fmt(sum(sp[s][p].values()))
rows.append(("Total", tot("before_balancing", "train"), tot("after_balancing", "train"), tot("before_balancing", "val"),
             tot("after_balancing", "val"), tot("after_balancing", "test")))
table(["Class", "Train (70 %)", "Train balanced", "Val (15 %)", "Val balanced", "Test (15 %, realistic)"], rows,
      widths=[1.1, 1.1, 1.1, 1.0, 1.0, 1.2], caption_text="Stratified split counts before and after balancing.")
n_mut, n_fix = len(feat["mutable_features"]), len(feat["fixed_features"])
para(f"After dropping identifiers (Flow ID, IP addresses, timestamp) and constant columns, {proj['n_features']} CICFlowMeter "
     f"features remained (final shapes: X_train {prep['final_shapes']['X_train'][0]:,}×59, X_val "
     f"{prep['final_shapes']['X_val'][0]:,}×59, X_test {prep['final_shapes']['X_test'][0]:,}×59). A StandardScaler was fitted "
     f"on the training set only. Features were partitioned into {n_mut} mutable features (duration, packet counts, byte totals, "
     "packet-length statistics, rates, inter-arrival times, header lengths, active/idle times) that the generator may change, and "
     f"{n_fix} fixed identity features that must be copied unchanged: {', '.join(feat['fixed_features'])}.")

doc.add_heading("8.2 Stage 2: Baseline IDS models", level=3)
table(["Model", "Architecture as implemented (code/ids_common.py)"], [
    ("Random Forest", "150 trees, default scikit-learn settings, all cores"),
    ("MLP", "Dense 128-64-32 with ReLU and dropout 0.3, linear output (4 classes)"),
    ("1D-CNN", "Conv1D(32, k=3)+MaxPool → Conv1D(64, k=3)+MaxPool → Dense 64 (dropout 0.3) → output"),
    ("LSTM", "Single LSTM layer, 64 hidden units, features fed as a 59-step sequence, dropout 0.2 → output")],
    widths=[1.3, 5.2], caption_text="Model configurations.")
para("All DL models were trained in PyTorch with Adam (Kingma and Ba, 2015, learning rate 1e-3), class-weighted cross-entropy, "
     "batch size 512, up to 40 epochs with early stopping on validation loss (patience 5). Each model was trained under three "
     "seeds (42, 102, 2025) with deterministic cuDNN settings, and all metrics are reported as mean ± standard deviation.")

doc.add_heading("8.3 Stage 3: LLM-based polymorphic variant generation", level=3)
para(f"We sampled {gen['llm']['source_rows_per_class']} source rows per attack class from the training split and sent them to "
     f"an LLM in batches of {gen['llm']['rows_per_call']} rows (temperature {gen['llm']['temperature']}, prompt version "
     f"{gen['llm']['prompt_version']}). The system prompt restricted the model to numeric tabular output, required fixed "
     "features to be copied exactly, required mutable features to stay inside the per-class [min, max] range, and required "
     "internal consistency. Each request applied one of four behavioural profiles:")
table(["Profile", "Instruction given to the LLM (summary)"], [
    ("extended_duration_low_rate", "Stretch Flow Duration ×1.5-4 and scale IAT/Active/Idle similarly; keep counts and bytes within ±20 %; recompute all per-second rates (they drop)."),
    ("high_variance_clustered", "Burstier traffic: raise IAT Std/Max ×1.5-3 while keeping IAT Min small; vary packet counts by 30-60 %; update bytes, headers and rates consistently."),
    ("reduced_magnitude", "Smaller packets: scale packet-length Max/Mean/Std by 0.4-0.8 (Min ≤ Mean ≤ Max), counts ×1.0-1.5; recompute byte totals and rates."),
    ("mixed_shift", "Moderate combined shift: timing, counts and sizes each scaled by an independent factor in [0.6, 1.6], then all totals and rates recomputed.")],
    widths=[1.9, 4.6], font=9, caption_text="The four behavioural profiles used for polymorphic generation.")
sc = gen["scripted"]
para(f"As a control (E3), a scripted generator produced {sc['variants_per_source']} variants per source row by multiplying "
     f"semantic feature groups by a shared log-normal factor (σ = {sc['group_sigma']}) plus per-feature noise "
     f"(σ = {sc['feature_sigma']}, maximum scale ×{sc['max_scale']}). After validation, the scripted pool was sub-sampled to "
     "match the LLM pool class by class, giving 685 class-matched scripted baseline variants.")

doc.add_heading("8.4 Five-step automated validation", level=3)
para("Every generated row, LLM or scripted, passed the same five checks; rows failing any check were discarded and counted:")
bullets([("Check 1 — Format/schema: ", "correct header, column count and numeric-only values."),
         ("Check 2 — Range clipping: ", "each mutable value clipped to the per-class [min, max] observed in training data."),
         ("Check 3 — Logical consistency: ", "non-negative counts, Min ≤ Mean ≤ Max for length and IAT statistics, and rates consistent with duration and totals."),
         ("Check 4 — Fixed-feature integrity: ", f"all {n_fix} identity features identical to the source row."),
         ("Check 5 — Copies/duplicates: ", "rows identical to their source or to another variant were removed.")], style="List Number")
v_rows = []
for src, label in (("llm", "LLM variants (E2)"), ("scripted", "Scripted variants (E3, before matching)")):
    r = val[src]
    v_rows.append((label, fmt(int(r["input_rows"])), r["Check 1 (format/schema)"], r["Check 2 (range clipping)"],
                   r["Check 3 (logical consistency)"], r["Check 4 (fixed-feature integrity)"], r["Check 5 (copies/duplicates)"],
                   f"{r['discarded']} ({float(r['discard_rate_%']):.1f} %)", fmt(int(r["kept_after_validation"]))))
table(["Source", "Input", "C1", "C2", "C3", "C4", "C5", "Discarded", "Kept"], v_rows,
      widths=[1.7, 0.6, 0.4, 0.4, 0.4, 0.4, 0.4, 1.1, 0.6], font=9,
      caption_text="Rows discarded at each validation check (03_validation_summary.csv).")
para(f"The LLM produced {n_llm_in} candidate rows; {n_llm_drop} ({float(val['llm']['discard_rate_%']):.1f} %) were discarded — "
     "36 for logical inconsistency, 1 for altering a fixed feature and 3 as copies — leaving 685 validated LLM variants "
     "(395 DoS_Hulk, 165 PortScan, 125 BruteForce). These were paired with 685 class-matched scripted variants for E3.")

doc.add_heading("8.5 Stage 4: Robustness evaluation", level=3)
para("The validated LLM pool was split, per class, into a training-variant pool (used only in Stage 5) and a held-out variant "
     "test set that was never used for training. Each baseline model was evaluated on the original test set (E1), on the "
     "held-out LLM variants (E2) and on the scripted variants (E3). Because the variant classes are imbalanced, the recall "
     "drop was computed in a class-balanced way (mean of per-class recalls).")
doc.add_heading("8.6 Stage 5: Adversarial retraining", level=3)
para("We added 20 % and 50 % of the training-variant pool to the original training set, retrained every model with identical "
     "settings and seeds, and evaluated it on both the original test set and the held-out variants (E4).")

# ----------------------------------------------------------------------------- 9 Metrics & design
doc.add_heading("9. Evaluation Metrics and Experimental Design", level=2)
table(["Metric", "Definition", "Role"], [
    ("F1-macro", "Unweighted mean of per-class F1 on the original test set", "Overall clean performance"),
    ("FPR", "Benign flows flagged as any attack", "Operational false-alarm cost"),
    ("Attack recall", "Attack flows detected as an attack class", "Main security metric"),
    ("Recall drop", "Clean attack recall − recall on variants", "Robustness (RQ1, RQ3)"),
    ("Evasion rate", "Share of variants classified as Benign", "Attack success (RQ2)"),
    ("Recovery", "Held-out variant recall after − before retraining", "Resilience (RQ4)"),
    ("ΔF1 / ΔFPR (orig)", "Change on the original test set after retraining", "Cost of the defense (RQ4)")],
    widths=[1.4, 3.3, 1.8], caption_text="Evaluation metrics.")
table(["Experiment", "Training data", "Test data", "Question"], [
    ("E1 Baseline", "Original train", "Original test", "Clean performance"),
    ("E2 Robustness", "Original train", "Held-out LLM variants", "RQ1, RQ2, RQ3"),
    ("E3 Control", "Original train", "Class-matched scripted variants", "Is the LLM stronger than noise?"),
    ("E4 Resilience", "Original train + 20 % / 50 % of variant pool", "Original test + held-out variants", "RQ4")],
    widths=[1.2, 2.0, 1.9, 1.4], caption_text="Experimental design (all experiments × 4 models × 3 seeds).")

# ----------------------------------------------------------------------------- 10 Results
doc.add_heading("10. Experimental Results and Discussion", level=2)
doc.add_heading("10.1 Summary table", level=3)
models = ["RandomForest", "MLP", "CNN1D", "LSTM"]
disp = {"RandomForest": "Random Forest", "MLP": "MLP", "CNN1D": "1D-CNN", "LSTM": "LSTM"}
cols1 = ["E1 F1-macro", "E1 FPR", "E1 attack recall", "E2 recall", "E2 recall drop", "E2 evasion", "E3 recall drop", "E3 evasion"]
table(["Model", "E1 F1-macro", "E1 FPR", "E1 attack recall", "E2 recall", "E2 recall drop", "E2 evasion", "E3 recall drop", "E3 evasion"],
      [[disp[mm]] + [S[mm][c] for c in cols1] for mm in models],
      widths=[0.9] + [0.7] * 8, font=8,
      caption_text="Clean, robustness and control results (mean ± std over seeds 42, 102, 2025; FINAL_summary_table.csv).")
cols2 = ["E4@20% held-out recall", "E4@20% recovery", "E4@20% ΔF1 (orig)", "E4@20% ΔFPR (orig)",
         "E4@50% held-out recall", "E4@50% recovery", "E4@50% ΔF1 (orig)", "E4@50% ΔFPR (orig)"]
table(["Model", "@20 % recall", "@20 % recovery", "@20 % ΔF1", "@20 % ΔFPR", "@50 % recall", "@50 % recovery", "@50 % ΔF1", "@50 % ΔFPR"],
      [[disp[mm]] + [S[mm][c] for c in cols2] for mm in models],
      widths=[0.9] + [0.7] * 8, font=8,
      caption_text="Adversarial retraining results E4 at 20 % and 50 % augmentation (held-out variants and original test set).")

doc.add_heading("10.2 E1 — Clean baseline", level=3)
para(f"On the original test set all four models performed well, confirming that the pipeline reproduces the high benchmark "
     f"scores reported in the literature. The Random Forest achieved the best clean performance (F1-macro "
     f"{m('RandomForest', 'E1 F1-macro'):.3f}, attack recall {m('RandomForest', 'E1 attack recall'):.3f}, FPR "
     f"{m('RandomForest', 'E1 FPR'):.4f}), followed by the 1D-CNN (F1 {m('CNN1D', 'E1 F1-macro'):.3f}), the MLP "
     f"(F1 {m('MLP', 'E1 F1-macro'):.3f}) and the LSTM (F1 {m('LSTM', 'E1 F1-macro'):.3f}). The LSTM was also the least "
     f"stable across seeds (FPR {S['LSTM']['E1 FPR']}), consistent with the fact that the 59 flow statistics have no natural "
     "temporal order for a recurrent model to exploit.")

figure(FIG / "fig1_original_vs_variants.png",
       "Attack recall of each model on the original test set (E1) versus held-out LLM-generated polymorphic variants (E2) "
       "and class-matched scripted variants (E3) (left), and macro-F1 on the original test set (right). Mean over three seeds; error bars show one standard deviation.")

doc.add_heading("10.3 E2 vs. E3 — Robustness under distribution shift", level=3)
para("The ranking by clean accuracy was almost the reverse of the ranking by robustness (Figure 1). The Random Forest, the "
     f"best model on clean data, suffered the sharpest drop on LLM variants: recall fell from {m('RandomForest', 'E1 attack recall'):.3f} "
     f"to {m('RandomForest', 'E2 recall'):.3f} (−{m('RandomForest', 'E2 recall drop'):.3f}), and {m('RandomForest', 'E2 evasion')*100:.1f} % "
     f"of the variants were classified as Benign. The LSTM followed (−{m('LSTM', 'E2 recall drop'):.3f}, "
     f"{m('LSTM', 'E2 evasion')*100:.1f} % evasion) with large seed-to-seed variance. In contrast, the MLP "
     f"(−{m('MLP', 'E2 recall drop'):.3f}, {m('MLP', 'E2 evasion')*100:.1f} % evasion) and the 1D-CNN "
     f"(−{m('CNN1D', 'E2 recall drop'):.3f}, {m('CNN1D', 'E2 evasion')*100:.1f} % evasion) proved substantially more resilient.")
para("We attribute this gap to the shape of the learned decision boundaries. A Random Forest partitions feature space with "
     "axis-aligned thresholds learned from the training distribution; when a profile such as extended_duration_low_rate "
     "stretches duration and lowers every per-second rate, a flow can cross several split thresholds at once and fall into "
     "leaves dominated by Benign traffic. The MLP and 1D-CNN learn smooth, distributed boundaries over standardised inputs and "
     "rely on many correlated features jointly, so a coherent multi-feature shift moves the logits gradually rather than "
     "flipping the decision.")
para("The scripted control confirmed that the LLM is a genuinely stronger adversary than random noise. LLM variants evaded every "
     f"model more often than class-matched scripted variants: {m('RandomForest', 'E2 evasion')*100:.1f} % vs. "
     f"{m('RandomForest', 'E3 evasion')*100:.1f} % on the Random Forest, {m('LSTM', 'E2 evasion')*100:.1f} % vs. "
     f"{m('LSTM', 'E3 evasion')*100:.1f} % on the LSTM, {m('CNN1D', 'E2 evasion')*100:.1f} % vs. {m('CNN1D', 'E3 evasion')*100:.1f} % "
     f"on the 1D-CNN and {m('MLP', 'E2 evasion')*100:.1f} % vs. {m('MLP', 'E3 evasion')*100:.1f} % on the MLP. Averaged over "
     "models, the LLM recall drop exceeded the scripted drop by +0.035. The LLM's advantage comes from producing coherent, "
     "internally consistent shifts in a chosen behavioural direction, whereas scripted noise scatters values around the source.")

figure(FIG / "fig2_evasion_rate.png",
       "Evasion rate (share of variants classified as Benign) per attack type and model for held-out LLM variants (E2, left) and scripted noise (E3, right), mean ± std over three seeds.")
para("Figure 2 breaks evasion down by attack type (RQ2). Averaged over models, PortScan variants were the most evasive "
     "(18.3 %), followed by DoS_Hulk (9.0 %) and BruteForce (6.4 %). PortScan flows are short and small, so modest changes "
     "to duration, packet counts and sizes move them close to ordinary short benign connections. BruteForce flows remain "
     "anchored by their fixed destination ports (21/22) and characteristic flag and window values, which the profiles were not "
     "allowed to change.")

figure(FIG / "fig3_before_after_retraining.png",
       "Effect of adversarial retraining with 0 %, 20 % and 50 % of the training-variant pool: held-out variant recall (left), "
       "macro-F1 (centre) and FPR (right) on the original test set, mean ± std over three seeds.")
doc.add_heading("10.4 E4 — Adversarial retraining at 20 % and 50 %", level=3)
rf_rec = m("RandomForest", "E4@20% recovery")
rf_drop = m("RandomForest", "E2 recall drop")
para(f"Adversarial retraining was highly effective for the model that needed it most (Figure 3). Adding just 20 % of the LLM "
     f"variant pool restored Random Forest held-out recall from {m('RandomForest', 'E2 recall'):.3f} to "
     f"{m('RandomForest', 'E4@20% held-out recall'):.3f} (+{rf_rec:.3f}), recovering {rf_rec / rf_drop * 100:.1f} % of the "
     f"recall it had lost, with zero clean-traffic penalty (ΔF1 = {m('RandomForest', 'E4@20% ΔF1 (orig)'):+.4f}, ΔFPR = "
     f"{abs(m('RandomForest', 'E4@20% ΔFPR (orig)')):.4f}). Increasing the share to 50 % gave no further gain "
     f"({m('RandomForest', 'E4@50% held-out recall'):.3f}), indicating that a small number of examples was enough for the trees "
     "to learn the new regions of feature space.")
para(f"The MLP and 1D-CNN, which lost little recall in E2, had little to recover: the MLP reached "
     f"{m('MLP', 'E4@50% held-out recall'):.3f} at 50 % augmentation and the 1D-CNN reached {m('CNN1D', 'E4@20% held-out recall'):.3f} "
     "at both levels, with changes on the original test set within ±0.004 F1 and ±0.001 FPR. The LSTM was the hardest model "
     f"to repair: at 20 % its held-out recall did not improve ({m('LSTM', 'E4@20% recovery'):+.3f}), and only 50 % augmentation "
     f"yielded a gain of +{m('LSTM', 'E4@50% recovery'):.3f}, reaching {m('LSTM', 'E4@50% held-out recall'):.3f} recall, still "
     "well below the other models. The LSTM's clean F1 did improve slightly (+0.015 to +0.022), but its high variance across "
     "seeds suggests that its sequential inductive bias is a poor fit for unordered flow statistics.")

doc.add_heading("10.5 Research questions summary", level=3)
table(["RQ", "Finding", "Answer"], [
    ("RQ1 — Recall drop", "Class-balanced recall drop on LLM variants: RF 0.207, LSTM 0.164, 1D-CNN 0.035, MLP 0.025.",
     "Yes — every model degraded; tree and recurrent models severely."),
    ("RQ2 — Evasion by attack", "Mean evasion over models: PortScan 18.3 %, DoS_Hulk 9.0 %, BruteForce 6.4 %.",
     "Evasion differs strongly by attack type; PortScan is most exposed."),
    ("RQ3 — Model comparison", "Most robust: MLP; least robust: Random Forest. LLM drop exceeds scripted-noise drop by +0.035 on average.",
     "Smooth neural models (MLP, 1D-CNN) were far more robust than RF and LSTM."),
    ("RQ4 — Retraining", "@20 %: RF +0.190, 1D-CNN +0.016, MLP +0.004, LSTM −0.013. @50 %: RF +0.188, LSTM +0.063, MLP +0.015, 1D-CNN +0.016. Clean ΔF1 ≤ |0.022|, ΔFPR ≤ |0.006|.",
     "Yes — retraining recovers most lost recall at negligible clean cost; LSTM needs more data.")],
    widths=[1.3, 3.3, 1.9], font=9, caption_text="Summary of answers to RQ1-RQ4 (FINAL_research_questions.txt).")

# ----------------------------------------------------------------------------- 11 Pipeline
doc.add_heading("11. Implementation and Reproducible Pipeline", level=2)
para("The study was delivered as five sequential notebooks backed by a shared module. All seeds, splits, feature lists, "
     "prompts and generation settings are stored in config/, and all metrics are written to results/ so every table and "
     "figure in this report can be regenerated.")
table(["Component", "What it does", "Main outputs"], [
    ("01_preprocessing.ipynb", "Load, clean, de-duplicate, split 70/15/15, balance, scale; define mutable/fixed features",
     "Processed arrays, scaler, feature_config.json, 01_preprocessing_report.json"),
    ("02_baseline_models.ipynb", "Train RF, MLP, 1D-CNN and LSTM under seeds 42, 102, 2025; E1 evaluation", "12 baseline models, E1 metrics"),
    ("03_llm_variant_generation.ipynb", "LLM generation with 4 profiles, scripted control, 5-step validation",
     "Variant pools, variant_generation_config.json, 03_validation_summary.csv"),
    ("04_robustness_evaluation.ipynb", "E2 and E3 evaluation, evasion per attack type", "Robustness metrics, Figures 1-2"),
    ("05_adversarial_retraining.ipynb", "E4 retraining at 20 % / 50 %, final tables", "FINAL_summary_table.csv, FINAL_research_questions.txt, Figure 3"),
    ("code/ids_common.py", "Shared models, training loop, early stopping, metrics, seeding", "Imported by all notebooks")],
    widths=[1.9, 2.6, 2.0], font=9, caption_text="The completed five-stage pipeline.")
para("Environment: Python 3, pandas, NumPy, scikit-learn 1.6 (Pedregosa et al., 2011), PyTorch and Matplotlib on Google Colab "
     "with a shared Google Drive project folder and a Git repository.")

# ----------------------------------------------------------------------------- 12 Ethics
doc.add_heading("12. Ethical and Responsible-Use Considerations", level=2)
bullets(["We worked only with feature-level numeric data from a public dataset; no real attack traffic, malware or exploit code "
         "was generated or used, and the LLM prompt explicitly forbade code, packets and payloads.",
         "Generated variants were used only to test and strengthen defensive models and are not published as an attack toolkit.",
         "Findings are presented in defensive terms, and the CIC-IDS2017 licence and citation requirements were respected."])

# ----------------------------------------------------------------------------- 13 Limitations
doc.add_heading("13. Limitations and Future Work", level=2)
doc.add_heading("13.1 Limitations", level=3)
bullets([("Small, imbalanced variant pool. ", "The final validated pool comprised 685 LLM variants with class imbalance "
          "(395 DoS_Hulk, 165 PortScan, 125 BruteForce), well below the original target of a few thousand per class, because of "
          "LLM generation throughput and safety-classifier constraints on attack-related requests. Per-class results for "
          "PortScan and BruteForce therefore rest on fewer samples."),
         ("Feature level, not packet level. ", "Evaluation operated on tabular CICFlowMeter statistical flow records rather "
          "than raw PCAP packet replay. Variants were checked for plausibility and internal consistency, but we did not prove "
          "that each corresponds to a working, protocol-valid attack."),
         ("Dataset dependence. ", "CIC-IDS2017 has known quality issues and a limited set of attacks; conclusions may not "
          "transfer to other environments."),
         ("Single LLM and prompt version. ", "Other LLMs or prompts may produce stronger or weaker variants."),
         ("Limited tuning. ", "Models were deliberately kept simple and lightly tuned; the LSTM in particular showed high seed variance.")])
doc.add_heading("13.2 Future work", level=3)
bullets(["Replay variants as packets in a sandboxed testbed and re-extract flows with CICFlowMeter to confirm end-to-end validity.",
         "Compare several LLMs and prompt strategies, and grow the pool to a balanced size per class.",
         "Repeat the study on UNSW-NB15 or CIC-IDS2018 to test generalisation.",
         "Sweep the augmentation share more finely (0, 5, 10, 20, 50 %) and add a variant/anomaly detector as a second defense layer.",
         "Test adaptive attackers who generate new variants against the retrained models."])

# ----------------------------------------------------------------------------- 14 Conclusion
doc.add_heading("14. Conclusion", level=2)
para("This project asked whether DL-based intrusion detection systems trained on a standard benchmark withstand attack traffic "
     "rewritten by an LLM. Across four models, three attack families and three seeds, the answer is nuanced. Clean benchmark "
     "accuracy was a poor predictor of robustness: the Random Forest, best on clean data, was the easiest to evade (22.5 %), "
     "while the MLP and 1D-CNN kept evasion below 3 %. LLM-generated variants were consistently more evasive than scripted "
     "noise, confirming that cheap generative tools give attackers a measurable advantage. A lightweight defense, adversarial "
     "retraining with a small share of validated variants, closed most of the gap at no cost on clean traffic.")
para("Security engineering recommendations:", bold_lead=None).runs[0].bold = True
bullets([("Do not select an IDS model on clean benchmark accuracy alone. ", "Add a robustness test on shifted or LLM-generated "
          "variants to the acceptance criteria, and report recall drop and evasion rate alongside F1."),
         ("Prefer smooth neural models (MLP, 1D-CNN) for flow-feature IDS, ", "or pair tree ensembles with a neural model, when "
          "robustness to behavioural drift matters."),
         ("Integrate LLM-driven red-teaming into the training loop. ", "Even 20 % augmentation with validated variants restored "
          "Random Forest recall to 0.982 with no false-alarm penalty; schedule this as periodic adversarial retraining."),
         ("Validate generated data rigorously. ", "A five-step schema, range, consistency, identity and duplicate check removed "
          "5.5 % of LLM rows and is essential before any synthetic data enters training."),
         ("Prioritise hardening for short, low-volume attacks such as port scans, ", "which proved the most evasive, and anchor "
          "detection on identity features (ports, flags, window sizes) that attackers cannot change cheaply."),
         ("Avoid sequence models for unordered flow statistics ", "unless the input has a real temporal structure; the LSTM was "
          "the least accurate, least stable and hardest to repair.")])

# ----------------------------------------------------------------------------- References
doc.add_heading("References", level=2)
for ref in [
    "Sharafaldin, I., Lashkari, A. H., and Ghorbani, A. A. (2018). Toward Generating a New Intrusion Detection Dataset and Intrusion Traffic Characterization. Proceedings of the 4th International Conference on Information Systems Security and Privacy (ICISSP).",
    "Moustafa, N., and Slay, J. (2015). UNSW-NB15: A Comprehensive Data Set for Network Intrusion Detection Systems. Military Communications and Information Systems Conference (MilCIS).",
    "Engelen, G., Rimmer, V., and Joosen, W. (2021). Troubleshooting an Intrusion Detection Dataset: the CICIDS2017 Case Study. IEEE Security and Privacy Workshops (SPW).",
    "Goodfellow, I. J., Shlens, J., and Szegedy, C. (2015). Explaining and Harnessing Adversarial Examples. International Conference on Learning Representations (ICLR).",
    "Kingma, D. P., and Ba, J. (2015). Adam: A Method for Stochastic Optimization. International Conference on Learning Representations (ICLR).",
    "Hochreiter, S., and Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation, 9(8).",
    "Breiman, L. (2001). Random Forests. Machine Learning, 45(1).",
    "Pedregosa, F., et al. (2011). Scikit-learn: Machine Learning in Python. Journal of Machine Learning Research, 12.",
]:
    p = doc.add_paragraph(ref, style="List Number")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

# page numbers in footer
fp = sec.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = fp.add_run()
for tag, text in (("begin", None), (None, "PAGE"), ("end", None)):
    if tag:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), tag)
    else:
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = text
    run._r.append(el)

doc.save(OUT)
print("saved", OUT)
