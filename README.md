# GitHub Pull Request Checker

A Python-based tool to analyze GitHub Pull Requests (PRs) in bulk from an Excel file. It checks merge status, authorship, and filters external comments — with results saved into a formatted Excel sheet.

------

## ✨ Features

- **Batch PR Check:** Process a list of PR URLs from Excel, retrieve their status, author, age, mergeability, and extract meaningful (non-bot) comments.
- **Fast & Asynchronous:** Powered by `asyncio` and `aiohttp` for high concurrency and speed.
- **Excel Report:** Exports all PR details into a new Excel file, auto-adjusting column widths for readability.
- **Terminal Summary Table:** Prints a neat summary statistics table after each run (pure ASCII, no extra packages needed).
- **Run History Management:** Automatically creates per-run timestamped folders under your Desktop and keeps only the latest N runs.
- **Progress Bar:** Real-time processing progress with the `rich` library.
- **Error Logging:** Per-run logs for diagnostics and traceability.
- **Flexible Filters:** Exclude system/bot users and unwanted keywords from comment collection.

------

## 🛠️ Prerequisites

- **Python 3.8+**
- **GitHub Personal Access Token**
   [Generate a token here](https://github.com/settings/tokens) (`repo` or `public_repo` scope is enough)

------

## 📥 Installation

Install dependencies with:

```bash
pip install openpyxl aiohttp rich
```

------

## ⚙️ Configuration

### 1. Set Your GitHub Token

Edit this line in the script to use your own token:

```python
GITHUB_TOKEN = "ghp_xxxxx..."  # Paste your token here!
```

*Keep this token secret and do not commit it to version control!*

### 2. Prepare the Input File

- Your input should be an Excel file (e.g., `PR_check.xlsx`)

- The **first column** must list valid GitHub PR URLs, like:

  | PR URL                                   |
  | ---------------------------------------- |
  | https://github.com/owner/repo/pull/123   |
  | https://github.com/another/repo/pull/456 |

  

- Update the script’s `INPUT_PATH` to the file location:

  ```python
  INPUT_PATH = r"C:\\Users\\your_name\\Desktop\\PR_check.xlsx"
  ```

### 3. (Optional) Adjust Script Parameters

- `API_DELAY`: Time delay (in seconds) between API requests (default: `0.1`)
- `CONCURRENT_REQUESTS`: Number of concurrent requests (default: `10`)
- `KEEP_LAST_N_RUNS`: Number of historical result folders to keep (default: `10`)

------

## 🚀 Usage

Simply run:

```bash
python github_pr_checker.py
```

During execution, you’ll see:

- Real-time progress bar
- Error and status updates
- Summary table at the end
- Excel result auto-opens on finish

------

## 📑 Output Explanation

### Excel Columns

| PR URL | Author | Merged Status | Open PRs > 1 week | External Comments Content |
| ------ | ------ | ------------- | ----------------- | ------------------------- |
| ...    | ...    | ...           | Yes / No          | ...                       |



- **Open PRs > 1 week:** Indicates PRs that are open, have no merge conflicts, and were created more than 7 days ago.

### Terminal Summary Table

A sample terminal output might look like this:

```sql
📊 Summary Statistics:

+--------------------------+-------+
| Metric                   | Value |
+--------------------------+-------+
| Total PRs                |    18 |
| Merged PRs               |     3 |
| Open PRs > 1 week        |     0 |
+--------------------------+-------+

⏱️ Execution time: 6.39 seconds
```

------

## 🔍 Column Details

- **Merged Status:** "Merged", "Not merged", or detailed state if blocked/conflicted.
- **Open PRs > 1 week:** Yes if PR is open, mergeable (no conflicts), and created over 7 days ago; No otherwise.
- **External Comments Content:** Displays only human (non-bot, non-system) comments. Automated/system comments are filtered out.

------

## 🔒 Security Notice

- **Do not share or commit your GitHub token!**
- Use a token with the minimum permissions needed.
- Be aware of [GitHub API rate limits](https://docs.github.com/en/rest/overview/resources-in-the-rest-api#rate-limiting).
   If you get rate limited, increase `API_DELAY` or lower `CONCURRENT_REQUESTS`.

------

## 🐞 Troubleshooting

- **"Not enough values to unpack":** Likely due to missing/invalid GitHub token, or an invalid PR URL in your Excel.
- **"Rate limit exceeded":** Increase `API_DELAY` or reduce `CONCURRENT_REQUESTS`.
- **Excel won't open:** Check if Excel is installed, or open the file manually from the results folder.

------

## 📄 License

This tool is provided for internal and automation use.
 No warranty is given for future API or format changes.

## 👨‍💻 Author

Created by Albert for internal PR audits. Contributions welcome if you wish to generalize this!


