---
name: worktree-start
description: Create a new isolated git worktree for parallel feature development.
---

# Worktree Start

Create a new git worktree and feature branch for parallel development.

Feature name: $ARGUMENTS

## Steps

1. Determine the repo name from the current directory name
2. Sanitize the feature name: replace spaces with hyphens, lowercase
3. Check if branch `feature/$ARGUMENTS` already exists:
   ```
   git branch --list "feature/$ARGUMENTS"
   ```
   If it exists, tell the user and suggest a different name. Stop.
4. Create the feature branch and worktree one level up from the current repo:
   ```
   git worktree add ../<repo-name>-$ARGUMENTS -b feature/$ARGUMENTS main
   ```
5. Verify creation with `git worktree list`
6. If a `.sln` or `.csproj` exists in the new worktree, verify it builds:
   ```
   powershell -ExecutionPolicy Bypass -File "<new-worktree-path>/scripts/build-quick.ps1"
   ```
7. Output this EXACT block so the user can copy-paste into a new terminal:

```
cd <full-absolute-path-to-new-worktree>
claude
```

8. Also list all active worktrees so the user can see what's running

## Important

- Use the FULL ABSOLUTE path (e.g. `C:\Users\tsmith\source\repos\swcsharpaddin-sheetmetal`), not relative
- Do NOT try to open a new terminal - just give the copy-paste block
- If the feature name contains spaces, replace them with hyphens
- If the branch already exists, tell the user and suggest a different name
