# Instructions for Claude Code in this folder

- `docs/` is the source of truth. Read `README.md`, then `docs/00` to `docs/18`, before changing anything.
- **Do not write production code in `app/` until the owner has signed off the design** (`docs/10-roadmap.md`,
  gate checklist). Throwaway scripts that verify facts are fine.
- `reference/prototype/` is reference only. Do not build on it.
- Every design change updates the affected doc; every decision is an ADR in `docs/decisions/`.
- Work on a branch, open a pull request, use conventional commit messages. Never push to `main`.
- Never commit secrets, `.env` files, MAC addresses, API keys, `*.db` or raw infrastructure output.
- This repo is **public**: never add a self-hosted runner, never use `pull_request_target`, pin third-party
  actions to a commit SHA, keep workflow permissions minimal.
- Keep shell scripts LF-only. Keep explanations short and direct.
