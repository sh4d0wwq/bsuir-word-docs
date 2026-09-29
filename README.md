# bsuir-word-docs

Cursor agent skill that creates Word documents per BSUIR standard СТП 01–2017: отчеты по лабораторным работам, пояснительные записки к курсовым и дипломным проектам.

## Install

```powershell
git clone https://github.com/sh4d0wwq/bsuir-word-docs "$HOME/.cursor/skills/bsuir-word-docs"
pip install python-docx
```

Requires the `word-document-server` MCP (`uvx --from office-word-mcp-server word_mcp_server`) and Microsoft Word on Windows for the `fix` step (updating contents, formulas, PDF export).

## Files

- `SKILL.md` — workflow, config and markup for the agent
- `reference.md` — writing rules from the standard
- `scripts/bsuir_docx.py` — `init` / `add` / `fix` builder
