# Contributing to Antigravity Excel Engine

First off, thank you for considering contributing to **Antigravity Excel Engine**! It's people like you who make open source such a great tool for the community.

## 🤝 Code of Conduct
By participating in this project, you agree to maintain a respectful, welcoming, and harassment-free environment for everyone.

## 🐛 How to Report a Bug
If you encounter a bug or unexpected behavior:
1. Ensure the bug was not already reported by searching on GitHub under [Issues](https://github.com/FranciscoGrandon/antigravity-excel-engine/issues).
2. If unable to find an open issue addressing the problem, open a new one using the **Bug Report** template.
3. Be sure to include:
   - Your OS version and Python version (`python --version`).
   - Excel version (Microsoft 365, 2021, etc.) and execution mode (`auto`, `live`, or `file`).
   - A minimal reproducible example (code snippet and sample workbook structure).
   - Full traceback logs if an exception occurred.

## 💡 Suggesting Enhancements
Feature requests are always welcome!
1. Check existing issues to verify whether your idea has already been proposed.
2. Open a new issue using the **Feature Request** template.
3. Explain why this enhancement would be useful to AI agents or Excel data pipelines, with concrete usage examples.

## 🛠️ Pull Request Process
1. Fork the repo and create your branch from `main`:
   ```bash
   git checkout -b feature/amazing-feature
   ```
2. Set up your local environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: .\venv\Scripts\activate
   pip install -r requirements.txt
   pip install pytest pytest-cov
   ```
3. Ensure your changes follow clean Python standards (PEP 8, type hints, informative docstrings).
4. Run the test suite:
   ```bash
   pytest tests/
   ```
5. Commit your changes with concise, descriptive commit messages:
   ```bash
   git commit -m "feat: add support for dynamic conditional formatting in set_cells"
   ```
6. Push to your fork and submit a Pull Request targeting `main`.

## 📜 License
By contributing to Antigravity Excel Engine, you agree that your contributions will be licensed under its [MIT License](LICENSE).
