# Editor HTML

Editor HTML otevírá soubory .html / .htm ve dvou režimech: **náhled** (vykreslená stránka) a **zdroj**.

- **Náhled**: skutečné vykreslování; relativní styly a obrázky se načítají vedle souboru.
- **Inspektor náhledu**: kliknutím vyberete prvek, dvojklikem upravíte jeho text přímo na místě, tlačítko na liště jej odstraní a Zeptat se AI se zeptá na výběr.
- **Režim zdroje**: úprava HTML; ctrl+F hledá, Nahradit vše uloží přepsané značky.
- **Ukládání**: věrné bajt po bajtu (BOM/CRLF/závěrečný řádek se zachovávají); uložení bez změn nic nepřepíše.
- **Přiblížení**: ctrl+kolečko / štípek mění měřítko náhledu; ctrl+Z uvnitř náhledu zruší poslední úpravu.

## Pruh nástrojů

Klikněte na libovolný prvek v náhledu a nad ním se objeví pruh nástrojů:

![Plovoucí pruh nástrojů nad vybraným prvkem](img/html-toolbar.png)

- **Soubor a historie**: Uložit, Uložit jako…, Zpět, Znovu, Najít; přepínač **Automatické ukládání** zapisuje změny v pravidelných intervalech.
- Přepínač **Náhled / Zdroj**; **Prezentovat** zobrazí stránku na celou obrazovku.
- **Formátování**: tučné, kurzíva, zvětšení a zmenšení velikosti písma; **panel stylů** pro vybraný prvek (barvy a další).
- **Vložit**: nadpis, odstavec, tabulka, obrázek (podle odkazu), další.
- **Operace s obrázkem** (s vybraným obrázkem): oříznout, **odebrat pozadí**, nahradit, zamknout poměr stran.
- **Operace s prvkem** (s vybraným prvkem v inspektoru náhledu): odstranit, duplikovat, posunout nahoru a dolů.
- **Tlačítko AI**: otevře panel AI; můžete se rovnou zeptat na vybraný prvek.

## Vložit kostru

Pro prázdnou stránku **Vložit ▸ Vložit kostru** zapíše minimální dokument ve standardním režimu:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <title></title>
  </head>
  <body></body>
</html>
```

Každá část tam je z nějakého důvodu, a proto je to příkaz, a ne něco, co se píše:

- **doctype**, jinak náhled běží v režimu quirks, kde se řízení velikosti rámečků a rozvržení tabulek řídí jinými pravidly, než čekáte;
- **`lang`**, jinak čtečka obrazovky nemá jazyk, ve kterém by stránku četla, a prohlížeč vybere písmo a kontrolu pravopisu pro nesprávný jazyk;
- **`charset`**, jinak se stránka s nelatinkovým textem může vykódovat jako mojibake.

Meta viewport je záměrně vynecháno: tento dokument se vykresluje v panelu na ploše bez mobilního viewportu, kterého by se to týkalo.

`lang` sleduje jazyk uživatelského rozhraní aplikace, takže kostru, kterou vložíte, je ta, na kterou jsou vaše nástroje již nastaveny. Potom ji můžete libovolně upravovat.

Položka se zobrazuje pouze v režimu úprav a pouze dokud je dokument prázdný — jakmile je tu obsah, není kam kostru *vložit*.

