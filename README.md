# Výpočet zákonného pojištění

Lokální webový prototyp pro přípravu výpočtu zákonného pojištění zaměstnavatele.

## Aktuální stav

- Rozhraní pro výběr screenshotu plateb, základního dokumentu a cílové šablony Excel.
- Kontrolní souhrn vybraných souborů.
- Soubory jsou zatím drženy pouze v paměti stránky; nejsou odesílány na server ani ukládány do historie.
- Automatické čtení PDF/Excel/Word, OCR, výpočty a zápis do cílového Excelu zatím nejsou implementovány. Rozhraní to výslovně uvádí, aby nedošlo k záměně prototypu za hotový výpočet.

## Ochrana dat

Stránka samotná neobsahuje síťová volání pro upload souborů. Pro skutečné nasazení je potřeba:
1. odstranit externí font import v `styles.css` nebo fonty lokálně přibalit;
2. přibalit knihovny pro čtení dokumentů a OCR jako lokální soubory, nikoli tahat skripty z CDN;
3. ověřit síťový provoz v prohlížeči a zajistit, aby obsah souborů, extrahovaný text i výsledky zůstávaly lokálně;
4. implementovat výpočty pouze podle ověřených pravidel z reálné šablony a dokumentů;
5. testovat proti ručně ověřeným příkladům před použitím pro skutečné pojistné.

GitHub Pages může sloužit k hostování statických souborů aplikace (dostupnost pro private repository závisí na tarifu a nastavení GitHubu). Samotný statický web nezaručuje soukromý přístup k aplikaci; data ale mohou zůstat na klientovi, pokud aplikace nic neodesílá.
