# ACIES-Tools Monorepo

Unified engineering toolchain integrating AutoCAD .NET plugins and Project Management orchestration.

## Structure

- **[`apps/ElectricalCommands`](apps/ElectricalCommands)**: AutoCAD C# / .NET solution containing custom drafting, circuiting, scheduling, and panel commands.
- **[`apps/ProjectManagement`](apps/ProjectManagement)**: Python + Web desktop application coordinating CAD plotting, scans, wire sizing, and project automation.
- **[`shared/`](shared/)**: Shared schemas, data contracts, and templates.
- **[`AGENTS.md`](AGENTS.md)**: AI agent context, architecture guidelines, and system instructions.

## Quick Start

### Build all projects
```powershell
.\build-all.ps1
```

### Open in VS Code / Antigravity
Open `ACIES-Tools.code-workspace` or the root folder `ACIES-Tools` in your editor.
