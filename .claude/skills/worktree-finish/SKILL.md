---
name: worktree-finish
description: Complete the current feature branch, commit, push, and provide cleanup instructions.
---

# Worktree Finish

Complete the current feature branch, commit, push, and provide merge/cleanup instructions.

## Steps

1. Check current branch name - confirm we're on a `feature/` branch (not main)
   ```
   git branch --show-current
   ```
   If on main, warn the user and stop.

2. Run the build to verify clean compile:
   ```
   powershell -ExecutionPolicy Bypass -File "scripts/build-quick.ps1"
   ```
   If build fails, stop and report the errors instead of completing.

3. Check for uncommitted changes with `git status`
   - If changes exist, stage and commit them with a descriptive message
   - Follow normal commit conventions (don't use --no-verify)

4. Push the branch:
   ```
   git push -u origin <branch-name>
   ```

5. List all worktrees: `git worktree list`

6. Determine the worktree folder name from the current working directory

7. Output this EXACT block so the user can copy-paste from their MAIN repo terminal to merge and clean up:

```
git merge --no-ff feature/<feature-name>
git worktree remove <full-absolute-path-to-this-worktree>
git branch -d feature/<feature-name>
```

8. Remind the user to close the terminal tab for this worktree

## Important

- Use FULL ABSOLUTE paths in all output
- Never force-delete branches - use `-d` not `-D`
- If build fails, stop and report the errors instead of completing
- Do NOT run the merge/cleanup commands yourself - the user must run them from the main worktree
