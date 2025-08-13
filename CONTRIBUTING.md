# Contributing to VBA DbContext

Thank you for considering contributing! We welcome contributions of all kinds: bug reports, feature requests, documentation improvements, and code changes.

## Table of Contents
- [Table of Contents](#table-of-contents)
- [Reporting Issues](#reporting-issues)
- [Suggesting Enhancements](#suggesting-enhancements)
- [Development Setup](#development-setup)
- [Coding Style](#coding-style)
- [Submitting Pull Requests](#submitting-pull-requests)
- [Code of Conduct](#code-of-conduct)

## Reporting Issues
If you encounter a bug or unexpected behavior, please open an issue in this repository:
1. Check existing issues to avoid duplicates.
2. Click **New issue**.
3. Provide a clear and descriptive title.
4. Describe steps to reproduce, expected vs. actual behavior, and any error messages.
5. Include minimal code samples or screenshots if applicable.

## Suggesting Enhancements
Feature requests and suggestions are welcome:
1. Open a new issue and choose the **feature request** template.
2. Explain your proposal, use cases, and possible API design.
3. Maintain backwards compatibility when possible.

## Development Setup
1. Fork this repository and clone your fork:
   ```bash
   git clone https://github.com/your-username/vba-dbcontext.git
   cd vba-dbcontext
   ```
2. Open your VBA host (Access, Excel, etc.) and import `DbContext.cls`.
3. In VBA, enable **Microsoft ActiveX Data Objects** reference.
4. You can test changes by running your own VBA scripts or sample database connections.

## Coding Style
- **Indentation**: Use 4 spaces per level.
- **Line Endings**: Windows CRLF.
- **Commenting**:  
  - English comments only.  
  - Use `’` (apostrophe) at line start.  
  - Provide summaries at procedure/module level.
- **Naming**:  
  - PascalCase for procedures and functions (e.g., `SelectQuery`).  
  - camelCase for local variables (e.g., `pConnection`).
- **Error Handling**: Use `On Error GoTo ErrorHandler` and set `LastError`.

## Submitting Pull Requests
1. Create a new branch for your change:
   ```bash
   git checkout -b feature/your-feature-name
   ```
2. Make your changes, commit with descriptive message:
   ```
   Add support for timeout parameter in Initialize()
   ```
3. Push your branch to your fork:
   ```bash
   git push origin feature/your-feature-name
   ```
4. Open a pull request against `main` branch of this repository.
5. In the PR description, reference related issues and describe your changes.
6. Wait for CI/documentation reviews and address feedback.

## Code of Conduct
This project follows the [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you agree to abide by its terms.

Thank you for helping make VBA DbContext better!
