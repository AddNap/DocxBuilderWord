
# DocxBuilderWord – Word Tiles Add-in (MVP)

Task pane do Word 365/2021 z kaflami (Jinja2), wczytywaniem placeholderów z dokumentu, wstawianiem w kursor oraz zablokowanym motywem.

## Struktura
```
.
├─ manifest-min.xml                  # sideload (Insert -> My Add-ins)
├─ manifest-with-shortcuts.xml       # wstążka + skrót Ctrl+Alt+=
└─ web/
   ├─ taskpane.html
   ├─ taskpane.js
   ├─ functions.html
   ├─ functions.js
   ├─ styles.css
   └─ assets/
      ├─ icon-16.png
      ├─ icon-32.png
      └─ icon-80.png
```

## Uruchomienie lokalne (HTTPS)
1. Postaw serwer **HTTPS** na `https://localhost:3000` dla folderu `web/` (np. `http-server` z certem dev lub `office-addin-dev-certs`).
2. W Wordzie: **Wstaw → Moje dodatki → Prześlij moje dodatki** i wskaż `manifest-min.xml` (lub `manifest-with-shortcuts.xml`).
3. W panelu wybierz/utwórz kafel i kliknij **+** aby dodać do listy, a następnie **Wstaw** albo główny **+**.

## Notatki
- „Globalny” klawisz `+` w dokumencie nie jest dostępny w Office.js; użyj przycisku **+** w panelu lub skrótu z manifestu.
- Znacznik `{{ INSERT_PRODUCT_TABLE }}` jest hakiem pod Twój renderer (np. `insert_product_table(...)`). 
- Bloki `{{ START_x }}` / `{{ END_x }}` obsłużysz np. w `conditional_block(...)`.
