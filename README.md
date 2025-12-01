# dietas-semanales-nbm
JSON de dietas semanales para Nutrition Blueprint Media™

## Generar el Excel d'inventari

El script `generate_inventory_excel.py` crea un llibre `inventari.xlsx` amb totes les
validacions, fórmules i el botó per exportar a JSON. Per fer-lo servir:

1. **Instal·la les dependències**
   ```bash
   python -m pip install --upgrade pip
   python -m pip install openpyxl
   ```
2. **Executa el generador** des del directori del projecte:
   ```bash
   python generate_inventory_excel.py
   ```
   Això crearà el fitxer `inventari.xlsx` a la mateixa carpeta.
3. **Obre l'Excel** i, quan ho demani, permet l'execució de macros.
4. **Edita les dades** a la pestanya «Inventari». Les fórmules i
   els desplegables s'apliquen automàticament.
5. **Exporta a JSON** prement el botó «Exportar a JSON»; es crearà un fitxer
   `inventory_export.json` al mateix directori.

> 💡 Si prefereixes treballar sense macros, pots ignorar el botó: la resta de
> funcionalitats del full funcionen igualment.
