# bsuir-word-docs

Cursor agent skill that creates Word documents per BSUIR standard СТП 01–2017: отчеты по лабораторным работам, пояснительные записки к курсовым и дипломным проектам.

Features: title page, automatic contents, Word-numbered headings (1 / 1.1 / 1.1.1) and lists («–», «1)», «а)»), formulas (Word equations), tables, figures, appendices, lint for missing references and placeholders, layout checks and page previews.

## Requirements

- Windows with Microsoft Word (the `fix` step updates contents, converts formulas, checks layout and exports PDF through Word)
- Python 3.11+
- [uv](https://docs.astral.sh/uv/) (to run the MCP server with `uvx`)

## Install

1. Clone the skill and install Python dependencies:

```powershell
git clone https://github.com/sh4d0wwq/bsuir-word-docs "$HOME/.cursor/skills/bsuir-word-docs"
pip install python-docx pymupdf
```

`pymupdf` is only needed for `preview` (page images).

2. Install uv if you do not have it:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

3. Add the [Office Word MCP Server](https://github.com/GongRzhe/Office-Word-MCP-Server) to `~/.cursor/mcp.json` (Cursor Settings → MCP → Add new global MCP server):

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "uvx",
      "args": ["--from", "office-word-mcp-server", "word_mcp_server"]
    }
  }
}
```

Keep the server name `word-document-server`: the skill refers to it by this name. Restart Cursor and check that the server is enabled (green) in Settings → MCP.

The MCP is used for point edits of an existing document; building from scratch works with the script alone.

## Usage

Ask the agent in Cursor, e.g. «Сделай отчет по лабораторной работе № 3 по СТП» and attach the materials (text, code files, images). The agent writes a config and markup chunks, then runs:

```powershell
python scripts/bsuir_docx.py build cfg.json OUT.docx ch1.md ch2.md
```

Unknown data (ФИО, группа, тема…) is left as `[[placeholders]]` and listed in the output.

## Files

- `SKILL.md` — workflow, config and markup for the agent
- `reference.md` — writing rules from the standard
- `scripts/bsuir_docx.py` — `init` / `add` / `fix` / `build` / `preview` builder
