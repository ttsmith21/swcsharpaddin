---
name: worktree-status
description: Show all active worktrees and their status.
---

# Worktree Status

Show all active git worktrees with their branch and change status.

## Steps

1. Run `git worktree list` and display the results

2. For each worktree, show:
   - Full absolute path
   - Branch name
   - Whether it has uncommitted changes:
     ```
     git -C <path> status --short
     ```

3. Output a copy-paste block for EACH worktree so the user can quickly jump to any one:

```
cd <full-absolute-path>
claude
```

4. If any worktrees have uncommitted changes, flag them with a warning

## Important

- Use FULL ABSOLUTE paths in all output
- If a worktree path is inaccessible, note it as potentially stale
- Show the main worktree separately from feature worktrees
