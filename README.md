
conversor_arquivos_formatos.py
Versão: 1.0

Este projeto é um simples exemplo de estudos e pesquisas, construído com Python e tkinter, para facilitar a conversão de arquivos entre os formatos PDF,(.docx) e (.xlsx).

⚙️ Requisitos: Windows, Microsoft Office instalado (para converter Word/Excel <-> PDF via COM)
Dependências pip: pywin32, pdf2docx, python-docx, pandas, openpyxl, tabula-py, pdfplumber, tqdm

Ele suporta conversões individuais e em lote para pastas inteiras.

✨ Recursos

* Interface Gráfica (GUI): Fácil de usar, baseada em tkinter.
* Conversões Múltiplas: Suporta as principais conversões entre PDF, Word e Excel.
*  Tenta usar `tabula-py` (mais robusto) para extração de tabelas de PDF para Excel e recorre a `pdfplumber` se necessário.
*  Utiliza `win32com.client` para conversões de Word/Excel para PDF, garantindo alta fidelidade de formato (requer o Office instalado).
* Converte todos os arquivos de um formato específico dentro de uma pasta.
