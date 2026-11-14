# Instructions for Claude Code in this folder

- `docs/` is the source of truth. Read `README.md` and `docs/00` onward before changing anything.
- Design first: no production code or firmware until the owner has chosen a direction and signed
  off the requirements. Throwaway scripts that verify facts about the device are fine.
- Do not flash the device without the owner's explicit go-ahead. Going back to the stock firmware
  may need a vendor activation code, so treat flashing as one-way.
- Every decision is an ADR in `docs/decisions/` (same template as `../trmnl-calendar`).
- Work on a branch, open a pull request, use conventional commit messages. Never push to `main`.
- This repo is **public**: never commit IP addresses, MAC addresses, Wi-Fi credentials, API keys,
  tokens, `.env` files or raw device dumps. Use placeholders such as `<screen-ip>`.
- Keep explanations short and direct.
