# Výpočet zákonného pojištění

Lokální webová aplikace pro načtení plateb, čtvrtletních předpisů, kontrolu rozpoznaných dat, výpočet nedoplatků a export výsledného Excelu.

## Vstupy

1. Screenshot se seznamem skutečných plateb (obrázek).
2. Základní dokument s předpisy a čtvrtletími: PDF, Excel, CSV, DOCX, TXT/HTML nebo obrázek.

## Aktuální funkce

- OCR screenshotů v prohlížeči pomocí Tesseract.js (čeština + angličtina).
- Čtení textové vrstvy PDF; u skenovaných PDF lokální OCR vykreslených stránek.
- Čtení XLSX/XLS/CSV, DOCX a textových/HTML souborů.
- Ruční oprava rozpoznaných plateb a čtvrtletních předpisů před výpočtem.
- Započítávání plateb nejstaršímu neuhrazenému řádnému pojistnému.
- Evidence navýšení za prodlení odděleně; platby se na navýšení použijí až po uhrazení řádného pojistného.
- Výpočet navýšení podle vzorců z dodané šablony: podíl 10 % za každý kalendářní měsíc mezi měsícem splatnosti a měsícem platby, pokud je platba po splatnosti; zaokrouhlení na celé Kč.
- Export sešitu se známým listem \`List1\`, přehledem čtvrtletí, rozpadem započtení, vstupními daty a kontrolními upozorněními.
- Kontrola rozporu mezi detekovanou sazbou a poměrem předpisů k výpočtovým základům.

## Ochrana dat

Aplikace neobsahuje API ani serverový endpoint pro upload dokumentů. Soubory se zpracovávají v aktuální relaci prohlížeče a automaticky se neukládají do GitHubu, databáze ani lokální historie. Exportovaný soubor se generuje v prohlížeči a stahuje do počítače.

Aby bylo možné načíst podporované formáty a provést OCR, aplikace stahuje JavaScriptové knihovny z CDNs a český/anglický OCR jazykový model. Tyto požadavky stahují pouze kód a jazyková data; aplikace jejich rozhraním nepředává obsah vstupních souborů. Pokud vyžadujete nulové požadavky na cizí domény, je třeba knihovny i OCR model přibalit lokálně a aktualizovat odkazy v \`index.html\` a \`app.js\`.

## Důležitá omezení

- OCR nemusí vždy správně přečíst malé částky nebo tabulky; před výpočtem zkontroluj a oprav hodnoty.
- U PDF a Word dokumentů závisí rozpoznání tabulek na jejich rozložení. Pokud parser nenajde všechny řádky, doplň je ručně.
- Starý formát DOC není podporován; převeď jej na DOCX nebo PDF.
- Před úředním použitím porovnej výsledky s ověřeným výpočtem. Aplikace má raději zobrazit varování než domýšlet chybějící pravidla nebo hodnoty.
- GitHub repo obsahuje pouze zdrojový kód; do repozitáře nevkládej screenshoty plateb, firemní dokumenty ani výstupní sešity.

## Spuštění

Aplikaci lze servírovat jako statické soubory. Pro lokální test stačí spustit jednoduchý HTTP server v adresáři aplikace a otevřít \`index.html\` přes \`localhost\`. Pokud ji budeš hostovat na GitHub Pages, pamatuj, že soukromé repo nezaručuje soukromý přístup na publikovanou webovou stránku.
