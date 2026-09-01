Secrets for local scripts live here:

1. Copy `env/.env.example` to `env/.env`
2. Put your Anthropic API key in `env/.env` as ANTHROPIC_API_KEY=...
3. Do not commit `env/.env` (it is listed in .gitignore if you use git).

Scripts load `env/.env` first, then fall back to a `.env` file in the project root if present.
