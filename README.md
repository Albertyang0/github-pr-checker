# GitHub Pull Request Checker

A Python-based tool to analyze GitHub Pull Requests (PRs) in bulk from an Excel file. It checks merge status, authorship, and filters external comments — with results saved into a formatted Excel sheet.

------

## 📌 Features

- ✅ Batch processing of PR URLs from Excel
- 🔍 Fetches merge status and mergeability info via GitHub API
- 👤 Extracts PR author and filters bot accounts
- 💬 Flags PRs with external human comments
- 📊 Auto-generates result Excel with color-coded statuses
- 📁 Maintains output history by date and timestamp
- 📈 Summary statistics at end of execution
- ⚙️ Async processing for high performance

------

## 📁 Input Requirements

An Excel file (e.g. `PR.xlsx`) with PR URLs in the **first column (A)**.

**Example:**

| PR URL                                 |
| -------------------------------------- |
| https://github.com/owner/repo/pull/123 |
| https://github.com/org/repo/pull/456   |



------

## 🧰 Prerequisites

Install required Python packages:

```bash
pip install aiohttp openpyxl rich
```

- Python version: **3.8+** recommended

------

## ⚙️ Configuration

Edit the variables at the top of the script:

```python
GITHUB_TOKEN = "your_github_token_here"  # Required
INPUT_PATH = r"C:\\Path\\To\\Your\\PR.xlsx"
```

### Optional Settings

| Variable              | Description                                    |
| --------------------- | ---------------------------------------------- |
| `API_DELAY`           | Delay between API calls (to avoid rate limits) |
| `CONCURRENT_REQUESTS` | Max concurrent GitHub API calls                |
| `KEEP_LAST_N_RUNS`    | Number of old result folders to retain         |



------

## 🚀 How to Run

Make sure your Excel file is ready, then execute the script:

```bash
python your_script_name.py
```

The program will:

1. Read PR URLs from the Excel file
2. Fetch status, author, and external comments for each PR
3. Write results to a new Excel file with extra columns
4. Display progress and generate a summary report

------

## 📤 Output

The tool creates a folder:

```swift
~/Desktop/PR_Check_Results/YYYY-MM-DD/Run_HHMMSS/
```

Inside, you'll find:

- `PR_Check_Result_Async.xlsx`: The annotated Excel file
- `PR_check_log.txt`: A text log of processed PRs

### Excel Columns

| Column                      | Description                                   |
| --------------------------- | --------------------------------------------- |
| `PR URL`                    | Original PR link                              |
| `Author`                    | GitHub username of PR author                  |
| `Merged Status`             | Status (Merged, Not merged, Blocked, etc.)    |
| `Has External Comment`      | Yes/No depending on presence of human comment |
| `External Comments Content` | Text of relevant comments                     |



------

## 📊 Summary Output

After completion, the script prints:

```yaml
📊 Summary Statistics:
🔢 Total PRs: 10
✅ Merged PRs: 7
💬 PRs with external comments: 3
```

------

## ❗ Notes & Limitations

- Only works with GitHub public or private repos (requires token)
- Assumes PR URLs are valid and well-formed
- Maximum ~10 comments per PR is assumed (no pagination)
- Make sure your token has `repo` access for private repos

------

## 🛡️ Security

Do **NOT** commit your GitHub token in code. Use environment variables or `.env` files for secure storage.

------

## 🧪 Example

```python
GITHUB_TOKEN = "ghp_your_token"
INPUT_PATH = r"C:\\Users\\me\\Desktop\\PR.xlsx"
```

Run and get an Excel file like:

| PR URL      | Author | Merged Status        | Has External Comment | External Comments Content |
| ----------- | ------ | -------------------- | -------------------- | ------------------------- |
| https://... | alice  | Merged               | No                   | None                      |
| https://... | bob    | Not merged (Blocked) | Yes                  | "Please update tests."    |



------

## 🧩 Ideas for Future Enhancements

- Slack/Teams notifications for blocked PRs
- GitHub Enterprise domain support
- Web dashboard summary
- CSV or Markdown export

------

## 👨‍💻 Author

Created by Albert for internal PR audits. Contributions welcome if you wish to generalize this!


