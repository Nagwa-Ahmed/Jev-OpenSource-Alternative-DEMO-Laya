# Laya Zero-Shot Classification Benchmark & Evaluation Suite

A modular, zero-shot text classification and evaluation tool built on top of **Laya** (a fast, lightweight alternative to standard zero-shot LLM routers and NLI models).

This project allows evaluating **any Hugging Face dataset** using a local Laya model checkpoint, supports all three Laya question types (**`choice`**, **`noul`**, and **`score`**), and automatically exports styled Excel reports detailing prediction accuracy and sample-by-sample classifications.

---

## Table of Contents

- [Project Structure](#project-structure)
- [Installation & Setup](#installation--setup)
- [How It Works](#how-it-works)
- [Question Types & Step-by-Step Usage](#question-types--step-by-step-usage)
  - [1. `choice` (Multi-Class Classification)](#1-choice-multi-class-classification)
  - [2. `noul` (Binary Yes/No Classification)](#2-noul-binary-yesno-classification)
  - [3. `score` (Ordinal Rating / Multi-Level Scale)](#3-score-ordinal-rating--multi-level-scale)
- [Using Custom Hugging Face Datasets](#using-custom-hugging-face-datasets)
- [Excel Export Format](#excel-export-format)
- [CLI Reference](#cli-reference)

---

## Project Structure

```text
├── laya_demo.py                  # Main execution script and evaluation pipeline
├── config.yaml                   # Default configuration file (Banking77 - choice)
├── requirements.txt              # Project dependencies
├── models/
│   └── laya/                     # Local model checkpoint directory
│       ├── model.safetensors     # Shipped model weights (~842 MB)
│       ├── rl_agent_config.json  # Laya model configuration
│       ├── encoder/              # Encoder configuration
│       └── tokenizer/            # Tokenizer configuration
├── configs/
│   ├── sst2_noul.yaml            # Binary classification example (noul)
│   └── sentiment_score.yaml      # Ordinal rating example (score)
└── .hf_cache/                    # Local Hugging Face cache (keeps C: drive clean)
```

---

## Installation & Setup

Follow these steps to set up an isolated Python environment and run the benchmark:

### 1. Prerequisites
- **Python 3.10+** installed on your system.
- Check your installation:
  ```bash
  python --version
  ```

### 2. Create a Virtual Environment (`venv`)

Creating a virtual environment ensures that the project's dependencies do not conflict with other Python packages on your system.

- **On Windows (Command Prompt / PowerShell):**
  ```powershell
  # Navigate to the project directory
  cd "d:\jev demo"

  # Create the virtual environment named '.venv'
  python -m venv .venv

  # Activate the virtual environment
  # In PowerShell:
  .\.venv\Scripts\Activate.ps1
  # (If you get an ExecutionPolicy error in PowerShell, run: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass)

  # OR in Command Prompt (cmd.exe):
  .venv\Scripts\activate.bat
  ```

- **On macOS / Linux:**
  ```bash
  # Navigate to the project directory
  cd /path/to/project

  # Create the virtual environment
  python3 -m venv .venv

  # Activate the virtual environment
  source .venv/bin/activate
  ```

> *Tip: Once activated, your terminal prompt will show `(.venv)` at the beginning of the line.*

### 3. Install Dependencies

With the virtual environment activated, install all required packages:
```bash
pip install -r requirements.txt
```

### 4. Automatic Model & Cache Management
- **Automatic Model Download:** The model weights (`convaiinnovations/laya`) will automatically download into [`models/laya/`](file:///d:/jev%20demo/models/laya) during the first run.
- **Disk Protection:** Dataset artifacts and Hugging Face caches are automatically redirected to the local `.hf_cache` folder inside the workspace, keeping your primary OS drive (`C:\`) clean and preventing out-of-disk crashes.

---

## How It Works

The evaluator executes the following lifecycle:
1. **Reads Config**: Loads YAML/JSON configuration specifying the dataset, sampling size, and task type.
2. **Loads Dataset**: Streams or loads the requested Hugging Face dataset and discovers class labels.
3. **Builds Task Schema**: Converts task settings into Laya's internal question format (`choice`, `noul`, or `score`).
4. **Local Zero-Shot Inference**: Runs Laya's `Router` on CPU using weights loaded directly from `./models/laya` (with safe memory paging).
5. **Generates Excel Report**: Computes the exact number of correct and incorrect predictions, calculates accuracy, and writes a two-sheet formatted workbook (`.xlsx`).

---

## Question Types & Step-by-Step Usage

Laya supports three question types. Each serves a different evaluation paradigm:

```
┌────────────────────────────────────────────────────────────────────────┐
│                          Laya Question Types                           │
├──────────────────┬─────────────────────────────┬───────────────────────┤
│ Type             │ Task Domain                 │ Output Format         │
├──────────────────┼─────────────────────────────┼───────────────────────┤
│ choice           │ Multi-class classification  │ Best option + probs   │
│ noul             │ Binary (True / False)       │ Probability of true   │
│ score            │ Ordinal scale (0 to N-1)    │ Expected float score  │
└──────────────────┴─────────────────────────────┴───────────────────────┘
```

---

### 1. `choice` (Multi-Class Classification)

#### What It Does
Evaluates multi-class classification tasks. Given a list of class labels or candidate categories, Laya assigns probabilities across all choices and selects the best matching candidate.

#### Configuration Format
In your YAML config:
```yaml
task:
  type: "choice"
  question_id: "intent"
  instructions: "Classify this banking customer query into the most relevant category."
  # If criteria is null, it is auto-built from top_k_classes or all classes.
  # Alternatively, provide a custom dict of {class_key: description}.
  criteria: null

evaluation:
  samples: 10          # Number of samples to test (null = all)
  top_k_classes: 5     # Select top-k most frequent classes in training set
```

#### Step-by-Step Example (Banking77)
Using [`config.yaml`](file:///d:/jev%20demo/config.yaml):
```bash
# Run with default config (top 5 classes, 10 samples)
python laya_demo.py --config config.yaml

# Override sample count and number of classes from CLI:
python laya_demo.py --config config.yaml --samples 25 --k 3
```

#### Output
- **Prediction:** The selected class key (e.g., `card_payment_fee_charged`).
- **Correctness:** Evaluated as `True Label == Predicted Label`.
- **Details:** Full probability distribution across all considered classes.

---

### 2. `noul` (Binary Yes/No Classification)

#### What It Does
`noul` evaluates binary questions. Rather than predicting categorical strings, Laya computes the calibrated scalar probability $P(\text{true}) \in [0.0, 1.0]$. If $P(\text{true}) \ge \text{threshold}$ (default `0.5`), the prediction is classified as `True`, otherwise `False`.

#### Configuration Format
In your YAML config:
```yaml
task:
  type: "noul"
  question_id: "sentiment"
  instructions: "Is this sentence expressing negative or critical sentiment?"
  positive_labels: [0, "0", "negative"]  # Dataset labels that map to True
  threshold: 0.5                         # Decision threshold
  criteria:
    false: "The sentence is positive or favorable"
    true: "The sentence is negative or critical"

evaluation:
  samples: 20
  top_k_classes: null                    # Not needed for binary tasks
```

#### Step-by-Step Example (SST-2 Sentiment)
Using [`configs/sst2_noul.yaml`](file:///d:/jev%20demo/configs/sst2_noul.yaml):
```bash
# Run binary sentiment classification on 50 samples
python laya_demo.py --config configs/sst2_noul.yaml --samples 50

# Output results to custom Excel filename
python laya_demo.py --config configs/sst2_noul.yaml --samples 50 --output sst2_binary_results.xlsx
```

#### Output
- **Prediction:** `True (p=0.9842)` or `False (p=0.0120)`.
- **Correctness:** Evaluated as `True Boolean == Predicted Boolean`.
- **Details:** Raw $P(\text{true})$ and the threshold used.

---

### 3. `score` (Ordinal Rating / Multi-Level Scale)

#### What It Does
`score` rates input text along an ordered scale (e.g., 0 to 4 Likert scale, severity ranking, or fine-grained sentiment). Laya returns:
1. An expected continuous score: $\sum_{i=0}^{N-1} i \cdot P(i)$.
2. Probabilities for each discrete level.
3. The most likely level ($\operatorname{argmax} P(i)$).

#### Configuration Format
In your YAML config:
```yaml
dataset:
  name: "talhaa/nlp-google-reviews-dataset"
  split: "train"
  train_split: "train"
  text_column: "review"
  label_column: "rating"    # 1 to 5 star ratings

  # Maps 1-5 star ratings to 0-4 discrete ordinal levels for Laya 'score'
  label_mapping:
    1: 0  # 1-star -> Level 0
    2: 1  # 2-star -> Level 1
    3: 2  # 3-star -> Level 2
    4: 3  # 4-star -> Level 3
    5: 4  # 5-star -> Level 4

task:
  type: "score"
  question_id: "review_rating"
  instructions: "Rate the customer sentiment and review rating on a scale from 0 (1-star, terrible) to 4 (5-star, excellent)."
  # Criteria MUST be an ordered list of level descriptions, index 0 first:
  criteria:
    - "1-star: terrible, strongly negative, major complaints, completely unsatisfied"
    - "2-star: poor, below expectations, negative experience with noticeable flaws"
    - "3-star: average, neutral or mixed feedback with both good and bad points"
    - "4-star: good, positive experience, pleased with the service or product"
    - "5-star: excellent, strongly positive, outstanding service, highly recommended"

evaluation:
  samples: 10
```

#### Step-by-Step Example (Google Reviews 1-5 Star Rating)
Using [`configs/sentiment_score.yaml`](file:///d:/jev%20demo/configs/sentiment_score.yaml):
```bash
# Run 5-point rating evaluation on Google Reviews (10 samples)
python laya_demo.py --config configs/sentiment_score.yaml --samples 10
```

#### Output
- **Prediction:** `Level 4 (5-star: excellent) [score=3.13]`.
- **Correctness:** Evaluated as `True Level Index == Predicted Level Index`.
- **Details:** Continuous expected rating score (e.g. `2.63`) and discrete level probabilities across all 5 star ratings.

---

## Using Custom Hugging Face Datasets

You can target any public dataset on the Hugging Face Hub by adjusting the `dataset` section of your configuration file:

```yaml
dataset:
  name: "rotten_tomatoes"     # Repository name on Hugging Face
  subset: null                # Optional configuration/subset name (if applicable)
  split: "test"               # Evaluation split ('test', 'validation', 'train')
  train_split: "train"        # Split used to discover label distributions
  text_column: "text"         # Name of the column with the input text
  label_column: "label"       # Name of the column with the target label
  
  # Optional explicit mapping if raw labels are numeric IDs:
  label_mapping:
    0: "negative"
    1: "positive"
```

---

## Excel Export Format

Every evaluation run creates an `.xlsx` workbook containing two formatted sheets:

### Sheet 1: `Summary`
Provides high-level metrics and test parameters:
- **Total Samples Evaluated**
- **Correct Count** (highlighted in green)
- **Incorrect Count** (highlighted in red)
- **Accuracy** ($XX.XX\%$)
- **Mean Confidence**
- Dataset name, split, task type, instructions, and run date.

### Sheet 2: `Detailed_Results`
Provides granular information for every evaluation sample:
- **Sample #**: Row index.
- **Input Text**: The text provided to the model.
- **True Label**: The ground truth label.
- **Predicted Label**: The model's prediction.
- **Result**: `CORRECT` (green) or `INCORRECT` (red).
- **Confidence**: Model confidence score ($0.0$ to $1.0$).
- **Raw Details / Probabilities**: Full dictionary of output probabilities or continuous scores.

---

## CLI Reference

The script accepts CLI arguments that override configuration file values on the fly:

| Flag | Description | Example |
|---|---|---|
| `--config` | Path to YAML or JSON config file (default: `config.yaml`) | `--config configs/sst2_noul.yaml` |
| `--samples` | Number of test samples to evaluate | `--samples 50` |
| `--k` | Top $k$ classes to filter for `choice` tasks | `--k 10` |
| `--task` | Override task type (`choice`, `noul`, `score`) | `--task noul` |
| `--output` | Destination path for the exported Excel report | `--output my_report.xlsx` |
| `--dataset` | Override Hugging Face dataset name | `--dataset mteb/tweet_sentiment_extraction` |

### Quick Run Commands

```bash
# 1. Multi-class banking query classification (top 5 classes, 10 samples)
python laya_demo.py --config config.yaml

# 2. Binary sentiment classification (50 samples)
python laya_demo.py --config configs/sst2_noul.yaml --samples 50

# 3. Google Reviews 1-5 star rating evaluation (10 samples)
python laya_demo.py --config configs/sentiment_score.yaml --samples 10

# 4. Custom output path
python laya_demo.py --config config.yaml --samples 30 --output D:/reports/banking_test.xlsx
```
