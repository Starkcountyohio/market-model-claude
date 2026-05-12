# Bethany — first-day setup

One-time setup to start working on this project. ~30 minutes if everything goes smoothly. If anything fails, send Jason a screenshot of the exact error message — don't try to fix it blind.

## 0. Before you start — confirm these two things with Jason

- **You have a GitHub account.** If not, sign up at https://github.com — free, takes 2 minutes. Use your `@starkcountyohio.gov` email so it's tied to your work identity. Tell Jason your GitHub username once you have one.
- **Jason has invited you to the repo.** You'll get an email from GitHub titled something like "@jmjeffries invited you to Starkcountyohio/market-model-claude". Click **Accept invitation**. After that, your GitHub account can see the repo.

You use YOUR OWN GitHub credentials throughout setup — never Jason's. VS Code will pop up a browser and ask you to sign in when you clone the repo (step 5); Windows caches the credentials after that, so you only sign in once per machine.

## 1. Accept the OneDrive share

Jason has shared the **Market Model Claude** folder with you in OneDrive. Open the email and click **Open**, then **Add shortcut to my OneDrive**. Wait for OneDrive to finish syncing (the green checkmark on the folder icon means done — initial sync is ~50 MB so should be quick).

Confirm the folder is at something like:
```
C:\Users\<your-username>\OneDrive - Stark County\Desktop\Market Model Claude\
```

Copy that path — you'll need it in step 4.

## 2. Verify prerequisites

You should already have these (Jason confirmed). Open PowerShell and run each to verify:

```powershell
python --version       # should print 3.11 or newer
git --version          # any 2.x
node --version         # any v18+
code --version         # any 1.8x+
```

If any are missing, install before continuing.

## 3. Install Python packages

In PowerShell:

```powershell
py -m pip install --upgrade pip
py -m pip install pandas numpy openpyxl lxml
```

(After cloning the repo in step 5, you can also run `py -m pip install -r requirements.txt` from the repo folder — same result.)

## 4. Set the data-path environment variable

This is the bridge between the git repo (code) and the OneDrive folder (data). Run ONCE in PowerShell, replacing the path with YOUR actual OneDrive path from step 1:

```powershell
[Environment]::SetEnvironmentVariable("MARKET_MODEL_DATA",
    "C:\Users\<your-username>\OneDrive - Stark County\Desktop\Market Model Claude",
    "User")
```

Then **close and reopen PowerShell** (and VS Code if it's open) so the new variable becomes visible. Verify:

```powershell
$env:MARKET_MODEL_DATA
# should print the path you set
```

## 5. Clone the git repo (in VS Code, not the terminal)

1. Open VS Code
2. `Ctrl+Shift+P` → type `Git: Clone` → Enter
3. Paste the repo URL Jason gave you (looks like `https://github.com/Starkcountyohio/market-model-claude.git`)
4. **Important:** when VS Code asks where to put it, pick `C:\Users\<your-username>\dev\` — NOT inside OneDrive. (If `dev` doesn't exist, create it.) OneDrive + git fight each other and corrupt repos.
5. When prompted "Would you like to open the cloned repository?" → click **Open**

VS Code may ask for GitHub credentials — sign in to your Starkcountyohio GitHub account.

## 6. Verify the bridge works

In VS Code, open the integrated terminal (Ctrl+`) and run:

```powershell
py -c "from config import SALES_CSV; print(SALES_CSV); print('exists:', SALES_CSV.exists())"
```

Expected output:
```
C:\Users\<your-username>\OneDrive - Stark County\Desktop\Market Model Claude\2026 Market Model\Sales Jan 2021-Dec 2025.CSV
exists: True
```

If it says `exists: False`, OneDrive hasn't finished syncing yet — wait a few minutes and re-run.

## 7. First session with Claude Code

1. In VS Code, open the **market-model-claude** folder (`File → Open Folder`)
2. Launch Claude Code from the side panel
3. Your first prompt: **"Read CLAUDE.md and tell me where the project stands and what my next step is on cluster 100."**

Claude will read the shared project context, summarize Track 1's findings, and orient you on the Track 2 work for cluster 100.

## 8. Daily workflow

Every session:
1. Open the repo folder in VS Code
2. Click the **Sync Changes** button at the bottom-left (round arrow icon). This pulls Jason's latest changes.
3. Work normally. Outputs you generate land in the OneDrive folder so Jason can see them.
4. When done:
   - **Source Control** panel (Ctrl+Shift+G) → check what files changed
   - Type a one-line message (e.g., "Cluster 100 — flagged FIREPLACE coefficient")
   - Click **Commit**, then **Sync Changes** to push

That's it. You'll never need git CLI commands.

## If something goes wrong

- **OneDrive shows the folder but files won't open** — right-click the folder → Always keep on this device
- **`MARKET_MODEL_DATA` env var "not set" error** — you forgot to restart PowerShell/VS Code after setting it
- **Sync Changes button greyed out** — there's nothing to pull or push; you're up to date
- **"Merge conflict"** — VS Code will mark the file; ask Jason to walk you through resolving it (rare if you stick to your clusters)
- **Anything else** — screenshot the error and send to Jason
