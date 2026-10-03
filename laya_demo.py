import os
import sys
import argparse
from collections import Counter
from datetime import datetime
import json
import yaml
import numpy as np
import pandas as pd

# ------------------------------------------------------------------------------
# 1. Environment & Console Setup
# ------------------------------------------------------------------------------
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Avoid filling up system C: drive by directing Hugging Face cache to workspace
_workspace_cache = os.path.abspath(".hf_cache")
os.environ.setdefault("HF_HOME", _workspace_cache)
os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

# Windows compatibility fix: safetensors mmap backend can trigger pagefile/access-violation
# crashes under low virtual memory or Windows paging limits. Using backend='pread' resolves this.
import safetensors.torch
_orig_load_file = safetensors.torch.load_file
def _safe_load_file(filename, device="cpu", backend="pread"):
    return _orig_load_file(filename, device=device, backend=backend)
safetensors.torch.load_file = _safe_load_file

from datasets import load_dataset
from huggingface_hub import snapshot_download
from laya import Router
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ------------------------------------------------------------------------------
# 2. Local Model Management
# ------------------------------------------------------------------------------
def ensure_local_model(model_dir: str) -> str:
    """Ensure the Laya model checkpoint exists in the specified local directory.
    If not found, downloads it directly to the local folder.
    """
    abs_model_dir = os.path.abspath(model_dir)
    safetensors_file = os.path.join(abs_model_dir, "model.safetensors")
    config_file = os.path.join(abs_model_dir, "rl_agent_config.json")

    if not (os.path.exists(safetensors_file) and os.path.exists(config_file)):
        print(f"Model checkpoint not found in '{abs_model_dir}'.")
        print("Downloading 'convaiinnovations/laya' directly to project directory...")
        snapshot_download(
            repo_id="convaiinnovations/laya",
            local_dir=abs_model_dir,
            allow_patterns=["rl_agent_config.json", "model.safetensors", "tokenizer/*", "encoder/*"],
        )
        print("Model downloaded successfully!")
    else:
        print(f"Using local model from: {abs_model_dir}")

    return abs_model_dir


# ------------------------------------------------------------------------------
# 3. Configuration Management
# ------------------------------------------------------------------------------
def load_config(config_path: str) -> dict:
    """Load configuration from YAML or JSON file."""
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        if config_path.endswith((".yaml", ".yml")):
            config = yaml.safe_load(f)
        elif config_path.endswith(".json"):
            config = json.load(f)
        else:
            # Try YAML parser as fallback for unspecified extensions
            config = yaml.safe_load(f)

    return config or {}


# ------------------------------------------------------------------------------
# 4. Dataset Helper Functions
# ------------------------------------------------------------------------------
def resolve_label_names(dataset, split: str, label_col: str, explicit_mapping: dict = None) -> dict:
    """Derive mapping from raw label ID to human-readable label string."""
    if explicit_mapping:
        # Convert keys to both int and str for robust lookup
        mapping = {}
        for k, v in explicit_mapping.items():
            mapping[k] = str(v)
            try:
                mapping[int(k)] = str(v)
            except (ValueError, TypeError):
                pass
        return mapping

    ds_split = dataset[split] if split in dataset else next(iter(dataset.values()))

    # Check datasets ClassLabel feature names
    if label_col in ds_split.features:
        feat = ds_split.features[label_col]
        if hasattr(feat, "names") and feat.names:
            return {i: name for i, name in enumerate(feat.names)}

    # Check companion column (e.g. label_text or label_name)
    companion_cols = [f"{label_col}_text", "label_text", "label_name"]
    for c_col in companion_cols:
        if c_col in ds_split.column_names:
            df_labels = ds_split.select_columns([label_col, c_col]).to_pandas().drop_duplicates()
            mapping = dict(zip(df_labels[label_col], df_labels[c_col]))
            # Also register string versions of keys
            for k, v in list(mapping.items()):
                mapping[str(k)] = str(v)
            return mapping

    # Fallback to unique values in the dataset
    unique_vals = list(set(ds_split[label_col]))
    return {v: str(v) for v in unique_vals}


# ------------------------------------------------------------------------------
# 5. Excel Exporter
# ------------------------------------------------------------------------------
def export_to_excel(output_path: str, summary_data: dict, detailed_results: list):
    """Create a formatted Excel spreadsheet with 'Summary' and 'Detailed_Results' sheets."""
    wb = openpyxl.Workbook()

    # Style definitions
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    title_font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    bold_font = Font(name="Calibri", size=11, bold=True)
    regular_font = Font(name="Calibri", size=11)
    
    correct_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    correct_font = Font(name="Calibri", size=11, bold=True, color="006100")
    
    incorrect_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    incorrect_font = Font(name="Calibri", size=11, bold=True, color="9C0006")

    thin_border_side = Side(style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # -----------------------------
    # Sheet 1: Summary
    # -----------------------------
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    # Title
    ws_summary["A1"] = "Laya Zero-Shot Evaluation Report"
    ws_summary["A1"].font = title_font
    ws_summary["A2"] = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws_summary["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

    # Table Header
    headers = ["Metric", "Value"]
    ws_summary.cell(row=4, column=1, value=headers[0]).fill = header_fill
    ws_summary.cell(row=4, column=1).font = header_font
    ws_summary.cell(row=4, column=2, value=headers[1]).fill = header_fill
    ws_summary.cell(row=4, column=2).font = header_font

    current_row = 5
    for metric, val in summary_data.items():
        cell_m = ws_summary.cell(row=current_row, column=1, value=metric)
        cell_v = ws_summary.cell(row=current_row, column=2, value=str(val))
        cell_m.font = bold_font
        cell_v.font = regular_font
        cell_m.border = thin_border
        cell_v.border = thin_border

        # Highlight key metrics
        if metric in ["Accuracy", "Correct Count"]:
            cell_v.font = correct_font
            cell_v.fill = correct_fill
        elif metric == "Incorrect Count" and int(val) > 0:
            cell_v.font = incorrect_font
            cell_v.fill = incorrect_fill

        current_row += 1

    # Auto-adjust column widths
    ws_summary.column_dimensions["A"].width = 32
    ws_summary.column_dimensions["B"].width = 60

    # -----------------------------
    # Sheet 2: Detailed Results
    # -----------------------------
    ws_details = wb.create_sheet(title="Detailed_Results")
    ws_details.views.sheetView[0].showGridLines = True

    detail_headers = [
        "Sample #",
        "Input Text",
        "True Label",
        "Predicted Label",
        "Result",
        "Confidence",
        "Raw Details / Probabilities"
    ]

    for col_idx, h_text in enumerate(detail_headers, start=1):
        c = ws_details.cell(row=1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, item in enumerate(detailed_results, start=2):
        ws_details.cell(row=row_idx, column=1, value=item["Sample"]).alignment = Alignment(horizontal="center")
        ws_details.cell(row=row_idx, column=2, value=item["Text"])
        ws_details.cell(row=row_idx, column=3, value=str(item["True Label"]))
        ws_details.cell(row=row_idx, column=4, value=str(item["Predicted Label"]))

        # Result status with conditional formatting
        res_cell = ws_details.cell(row=row_idx, column=5, value=item["Result"])
        res_cell.alignment = Alignment(horizontal="center")
        if item["Result"] == "CORRECT":
            res_cell.fill = correct_fill
            res_cell.font = correct_font
        else:
            res_cell.fill = incorrect_fill
            res_cell.font = incorrect_font

        conf_cell = ws_details.cell(row=row_idx, column=6, value=float(item["Confidence"]))
        conf_cell.alignment = Alignment(horizontal="right")
        conf_cell.number_format = "0.0000"

        ws_details.cell(row=row_idx, column=7, value=str(item.get("Details", "")))

        for col_idx in range(1, 8):
            ws_details.cell(row=row_idx, column=col_idx).border = thin_border
            if col_idx != 5:  # keep custom font for result
                ws_details.cell(row=row_idx, column=col_idx).font = regular_font

    # Set column widths for details sheet
    col_widths = {
        "A": 10,
        "B": 65,
        "C": 30,
        "D": 30,
        "E": 14,
        "F": 14,
        "G": 45
    }
    for col_letter, width in col_widths.items():
        ws_details.column_dimensions[col_letter].width = width

    # Save workbook
    abs_output = os.path.abspath(output_path)
    os.makedirs(os.path.dirname(abs_output) or ".", exist_ok=True)
    wb.save(abs_output)
    print(f"\n[OK] Excel report successfully exported to: {abs_output}")


# ------------------------------------------------------------------------------
# 6. Main Execution Pipeline
# ------------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Configurable Zero-Shot Dataset Evaluator with Laya.")
    parser.add_argument("--config", type=str, default="config.yaml", help="Path to YAML/JSON configuration file (default: config.yaml)")
    parser.add_argument("--samples", type=int, default=None, help="Override number of evaluation samples")
    parser.add_argument("--k", type=int, default=None, help="Override top-k classes (for choice task)")
    parser.add_argument("--task", type=str, default=None, choices=["choice", "noul", "score"], help="Override task type")
    parser.add_argument("--output", type=str, default=None, help="Override output Excel path")
    parser.add_argument("--dataset", type=str, default=None, help="Override Hugging Face dataset name")
    args = parser.parse_args()

    # Load configuration
    print(f"Loading configuration from: {args.config}")
    cfg = load_config(args.config)

    # Apply command-line overrides
    model_dir = cfg.get("model", {}).get("path", "models/laya")
    device = cfg.get("model", {}).get("device", "cpu")

    ds_cfg = cfg.get("dataset", {})
    ds_name = args.dataset or ds_cfg.get("name", "mteb/banking77")
    ds_subset = ds_cfg.get("subset", None)
    ds_split = ds_cfg.get("split", "test")
    ds_train_split = ds_cfg.get("train_split", "train")
    text_col = ds_cfg.get("text_column", "text")
    label_col = ds_cfg.get("label_column", "label")
    explicit_label_map = ds_cfg.get("label_mapping", None)

    eval_cfg = cfg.get("evaluation", {})
    num_samples = args.samples if args.samples is not None else eval_cfg.get("samples", 10)
    random_seed = eval_cfg.get("random_seed", 42)
    top_k = args.k if args.k is not None else eval_cfg.get("top_k_classes", None)

    task_cfg = cfg.get("task", {})
    task_type = args.task or task_cfg.get("type", "choice").lower()
    qid = task_cfg.get("question_id", "task_q")
    instructions = task_cfg.get("instructions", "Classify the input.")
    task_criteria = task_cfg.get("criteria", None)

    export_path = args.output or cfg.get("export", {}).get("excel_path", "laya_evaluation_results.xlsx")

    # 1. Ensure local model exists
    local_model_path = ensure_local_model(model_dir)

    # 2. Load Hugging Face dataset
    print(f"\nLoading dataset '{ds_name}' (subset={ds_subset})...")
    if ds_subset:
        dataset = load_dataset(ds_name, ds_subset)
    else:
        dataset = load_dataset(ds_name)

    if ds_split not in dataset:
        raise ValueError(f"Split '{ds_split}' not found in dataset. Available: {list(dataset.keys())}")

    eval_ds = dataset[ds_split]
    train_ds = dataset.get(ds_train_split, eval_ds)

    # Resolve label names
    label_map = resolve_label_names(dataset, ds_train_split, label_col, explicit_label_map)

    # 3. Configure Task Specifics
    top_k_class_names = []
    if task_type == "choice":
        if top_k and top_k > 0:
            print(f"Selecting top {top_k} classes from '{ds_train_split}' split...")
            train_labels = train_ds[label_col]
            label_counts = Counter(train_labels)
            top_k_labels = [label for label, _ in label_counts.most_common(top_k)]
            top_k_class_names = [label_map.get(lbl, label_map.get(str(lbl), str(lbl))) for lbl in top_k_labels]
            print(f"Top {top_k} classes:\n - " + "\n - ".join(top_k_class_names))

            # Filter evaluation dataset to top-k
            top_k_set = set(top_k_labels) | {str(l) for l in top_k_labels}
            eval_ds = eval_ds.filter(lambda x: x[label_col] in top_k_set)
        else:
            top_k_class_names = [label_map[k] for k in sorted(label_map.keys())]

        # Build criteria for choice
        if not task_criteria:
            criteria = {name: name.replace("_", " ") for name in top_k_class_names}
        else:
            criteria = task_criteria

        questions = {
            qid: {
                "type": "choice",
                "instructions": instructions,
                "criteria": criteria
            }
        }

    elif task_type == "noul":
        # Binary classification
        criteria = task_criteria
        positive_labels = task_cfg.get("positive_labels", [1, "positive", "true", "True"])
        pos_set = {str(x).lower().strip() for x in positive_labels}
        threshold = float(task_cfg.get("threshold", 0.5))

        q_def = {
            "type": "noul",
            "instructions": instructions
        }
        if criteria:
            q_def["criteria"] = criteria
        questions = {qid: q_def}

    elif task_type == "score":
        # Ordinal scale / ratings
        criteria = task_criteria
        if not criteria or not isinstance(criteria, list):
            raise ValueError("Task type 'score' requires 'criteria' as a list of level descriptions, index 0 first.")

        questions = {
            qid: {
                "type": "score",
                "instructions": instructions,
                "criteria": criteria
            }
        }
    else:
        raise ValueError(f"Unsupported task type '{task_type}'. Must be one of: 'choice', 'noul', 'score'.")

    # Sample evaluation dataset
    total_available = len(eval_ds)
    sample_count = min(num_samples, total_available) if num_samples else total_available
    sample_ds = eval_ds.shuffle(seed=random_seed).select(range(sample_count))

    # Initialize Laya Router with local model
    print(f"\nInitializing Laya Router with local model from: {local_model_path}...")
    router = Router(
        models={"english": local_model_path},
        device=device
    )

    # 4. Run Evaluation
    detailed_results = []
    correct_count = 0
    incorrect_count = 0

    print(f"\nRunning Laya zero-shot evaluation on {sample_count} samples (task={task_type})...\n")

    for i, item in enumerate(sample_ds):
        text = str(item[text_col])
        raw_label = item[label_col]

        # Predict
        prediction = router.predict(text, questions, model="english")
        ans = prediction["answers"][qid]
        confidence = float(ans["confidence"])

        # Determine correctness and format predictions based on task type
        if task_type == "choice":
            true_label_name = label_map.get(raw_label, label_map.get(str(raw_label), str(raw_label)))
            pred_choice = ans["choice"]
            is_correct = (str(true_label_name).strip().lower() == str(pred_choice).strip().lower())
            pred_display = str(pred_choice)
            true_display = str(true_label_name)
            details = json.dumps(ans.get("probabilities", {}))

        elif task_type == "noul":
            # True label to boolean
            raw_str = str(raw_label).lower().strip()
            label_text = str(label_map.get(raw_label, raw_label)).lower().strip()
            is_true_positive = (raw_str in pos_set) or (label_text in pos_set)
            true_display = "True" if is_true_positive else "False"

            noul_prob = float(ans["noul"])
            pred_bool = (noul_prob >= threshold)
            pred_display = f"{'True' if pred_bool else 'False'} (p={noul_prob:.4f})"
            is_correct = (pred_bool == is_true_positive)
            details = f"noul_prob={noul_prob:.4f}, threshold={threshold}"

        elif task_type == "score":
            score_val = float(ans["score"])
            probs = ans.get("probabilities", {})
            pred_level = int(np.argmax([probs.get(str(lvl), 0.0) for lvl in range(len(criteria))]))
            pred_display = f"Level {pred_level} ({criteria[pred_level]}) [score={score_val:.2f}]"

            # Check if raw_label matches level
            try:
                mapped_val = label_map.get(raw_label, label_map.get(str(raw_label), raw_label))
                true_level = int(mapped_val)
                # If dataset uses 1-based ratings (1..N) and criteria has N levels, adjust offset if needed
                if true_level == len(criteria) and 0 not in [int(k) for k in label_map.keys() if str(k).isdigit()]:
                    true_level -= 1

                if 0 <= true_level < len(criteria):
                    true_display = f"Level {true_level} [{raw_label}*] ({criteria[true_level]})"
                else:
                    true_display = f"Level {true_level} [{raw_label}*]"
                is_correct = (pred_level == true_level)
            except (ValueError, TypeError):
                true_display = str(raw_label)
                is_correct = False

            details = f"expected_score={score_val:.4f}, probs={json.dumps(probs)}"

        if is_correct:
            correct_count += 1
            result_str = "CORRECT"
        else:
            incorrect_count += 1
            result_str = "INCORRECT"

        mark = "[CORRECT]  " if is_correct else "[INCORRECT]"
        print(f"  [{i+1}/{sample_count}] {mark} conf={confidence:.3f} | True: '{true_display}' | Pred: '{pred_display}'")

        detailed_results.append({
            "Sample": i + 1,
            "Text": text[:120] + ("..." if len(text) > 120 else ""),
            "True Label": true_display,
            "Predicted Label": pred_display,
            "Result": result_str,
            "Confidence": confidence,
            "Details": details
        })

    # Summary Statistics
    accuracy = (correct_count / sample_count * 100.0) if sample_count > 0 else 0.0
    mean_conf = float(np.mean([r["Confidence"] for r in detailed_results])) if detailed_results else 0.0

    print("\n" + "=" * 60)
    print("EVALUATION SUMMARY")
    print("=" * 60)
    print(f"Dataset:            {ds_name} (split={ds_split})")
    print(f"Task Type:          {task_type}")
    print(f"Total Samples:      {sample_count}")
    print(f"Correct Count:      {correct_count}")
    print(f"Incorrect Count:    {incorrect_count}")
    print(f"Accuracy:           {accuracy:.2f}%")
    print(f"Mean Confidence:    {mean_conf:.4f}")
    print("=" * 60)

    summary_data = {
        "Dataset": ds_name,
        "Subset": str(ds_subset),
        "Evaluation Split": ds_split,
        "Task Type": task_type.upper(),
        "Total Samples Evaluated": sample_count,
        "Correct Count": correct_count,
        "Incorrect Count": incorrect_count,
        "Accuracy": f"{accuracy:.2f}%",
        "Mean Confidence": f"{mean_conf:.4f}",
        "Top K Classes Filter": top_k if (task_type == "choice" and top_k) else "All",
        "Instructions": instructions,
        "Random Seed": random_seed,
        "Model Path": local_model_path,
        "Execution Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # 5. Export to Excel
    export_to_excel(export_path, summary_data, detailed_results)


if __name__ == "__main__":
    main()
