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
- Dvacet průchodů OCR nad pěti různými variantami obrazu a čtyřmi režimy rozložení stránky; platby se konsolidují hlasováním po řádcích, ne porovnáním celého textu znak po znaku.
- Zvýrazněný tmavý řádek plateb se navíc čte samostatně s převráceným kontrastem, aby se neztratila první úhrada.
- U obrázků s mřížkou se hranice tabulky a sloupců hledají podle pixelů a OCR čte datum, základ a pojistné z oddělených buněk; číslice se nespojují přes hranice sloupců.
- Když jedna částka zjevně nesedí na společnou sazbu potvrzenou ostatními čtvrtletími, aplikace může opravit pravděpodobnou chybu OCR podle vzorce a zaznamená původní i opravenou hodnotu do auditního textu.
- Přesně 20 kontrolních bodů (shoda čtení podle sloupců, formát dat, návaznost čtvrtletí, věrohodnost základu, shoda pojistného se sazbou, duplicity a oddělení základ/pojistné).
- Neúspěšné kontroly jsou upozornění, nikoli blokace: výpočet i export lze spustit i při neshodě OCR nebo neúspěšném validačním testu. Chybějící povinné hodnoty mohou samozřejmě znemožnit samotný matematický výpočet.
- Samostatné sloupce pro požadovanou částku platby a skutečně vyřízenou částku; nulově vyřízené platby se nezapočítají, ale zůstanou ve vstupním auditu.
- Text OCR screenshotu plateb zůstává dostupný pro vizuální kontrolu.
- Započítávání plateb nejstaršímu neuhrazenému řádnému pojistnému.
- Evidence navýšení za prodlení odděleně; platby se na navýšení použijí až po uhrazení řádného pojistného.
- Výpočet navýšení podle vzorců z dodané šablony: podíl 10 % za každý kalendářní měsíc mezi měsícem splatnosti a měsícem platby, pokud je platba po splatnosti; zaokrouhlení na celé Kč.
- Export sešitu se známým listem \`List1\`, přehledem čtvrtletí, rozpadem započtení, vstupními daty a kontrolními upozorněními.
- Kontrola rozporu mezi detekovanou sazbou a poměrem předpisů k výpočtovým základům.

## Ochrana dat

Aplikace neobsahuje API ani serverový endpoint pro upload dokumentů. Soubory se zpracovávají v aktuální relaci prohlížeče a automaticky se neukládají do GitHubu, databáze ani lokální historie. Exportovaný soubor se generuje v prohlížeči a stahuje do počítače.

Aby bylo možné načíst podporované formáty a provést OCR, aplikace stahuje JavaScriptové knihovny z CDNs a český/anglický OCR jazykový model. Tyto požadavky stahují pouze kód a jazyková data; aplikace jejich rozhraním nepředává obsah vstupních souborů. Pokud vyžadujete nulové požadavky na cizí domény, je třeba knihovny i OCR model přibalit lokálně a aktualizovat odkazy v \`index.html\` a \`app.js\`.

## Důležitá omezení

- Ani 20 shodných OCR průchodů není samo o sobě matematický důkaz správnosti. Proto se současně kontroluje sazba, výpočtový základ, předpisy, datumy a návaznost čtvrtletí.
- Rozdíly mezi celými OCR výstupy nejsou samy o sobě důvodem k blokaci; aplikace používá hlasování po řádcích, čtení samostatných buněk a 20 validačních kontrol. Neúspěšné kontroly se zobrazí jako upozornění.
- OCR nemusí vždy správně přečíst malé částky nebo tabulky; před výpočtem zkontroluj a oprav hodnoty.
- U PDF a Word dokumentů závisí rozpoznání tabulek na jejich rozložení. Pokud parser nenajde všechny řádky, doplň je ručně.
- Starý formát DOC není podporován; převeď jej na DOCX nebo PDF.
- Před úředním použitím porovnej výsledky s ověřeným výpočtem. Aplikace má raději zobrazit varování než domýšlet chybějící pravidla nebo hodnoty.
- GitHub repo obsahuje pouze zdrojový kód; do repozitáře nevkládej screenshoty plateb, firemní dokumenty ani výstupní sešity.

## Spuštění

Aplikaci lze servírovat jako statické soubory. Pro lokální test stačí spustit jednoduchý HTTP server v adresáři aplikace a otevřít \`index.html\` přes \`localhost\`. Pokud ji budeš hostovat na GitHub Pages, pamatuj, že soukromé repo nezaručuje soukromý přístup na publikovanou webovou stránku.

- Oprava záložního parsování: pokud řádky nelze vytáhnout z OCR textu, předpisy se sestaví z rozpoznaných pozic buněk v tabulce.
