# PMax Asset Scheduler

**Időzített asset kezelés Performance Max kampányokhoz Google Ads-ban.**

Automatikusan hozzáad és eltávolít szöveges, kép és videó asset-eket az asset group-okból ütemezetten - ideális promóciókhoz, szezonális kampányokhoz.

---

## Tartalomjegyzék

- [Funkciók](#funkciók)
- [Telepítés](#telepítés)
- [Konfiguráció](#konfiguráció)
- [Sheet Struktúra](#sheet-struktúra)
  - [TextAssets](#textassets)
  - [ImageAssets](#imageassets)
  - [VideoAssets](#videoassets)
- [Használat](#használat)
- [Linkek és Thumbnailok](#linkek-és-thumbnailok)
- [Státuszok és Hibaüzenetek](#státuszok-és-hibaüzenetek)
- [Limitek](#limitek)
- [Gyakori Kérdések](#gyakori-kérdések)
- [Ismert Korlátozások](#ismert-korlátozások)
- [Verzió](#verzió)

---

## Funkciók

- **Időzített hozzáadás (ADD)**: Asset-ek automatikus hozzáadása megadott dátumon és órában
- **Időzített törlés (REMOVE)**: Asset-ek automatikus eltávolítása megadott dátumon és órában
- **Három asset típus támogatása**:
  - Szöveges asset-ek (headlines, long headlines, descriptions)
  - Kép asset-ek (horizontal, square, vertical)
  - Videó asset-ek (YouTube videók)
- **Validáció**: Előzetes ellenőrzés limitek, duplikációk és hibák ellen
- **Post-execution preview**: Execution után friss állapot alapján újravalidál és frissíti a preview sheet-et
- **Email értesítések**: Preview és Execution riportok
- **Post-verification**: Műveletek sikeres végrehajtásának ellenőrzése
- **Kép és videó linkek**: Kattintható thumbnail-ek a bemeneti sheet-eken (ImageAssets, VideoAssets); képeknél szöveges link a riport sheet-eken és az emailben is — videóknál csak a bemeneti sheet-en kattintható

---

## Telepítés

### 1. Google Sheet létrehozása

> ⚠️ **Frissítés régebbi verzióról?** Készíts új sheetet az alábbi sablonból (ne a régit módosítsd), és a meglévő `Results` fület töröld a sheetből — az oszlopszerkezet megváltozott.

Másold le a sablon sheet-et:
**[Sablon Sheet Másolása](https://docs.google.com/spreadsheets/d/14cWq7Ly8XHIkcUkIhYHeIwhwEb3w5Qjl8npcf6SwaiI/copy)**

### 2. Script hozzáadása

1. Nyisd meg a Google Ads fiókodat
2. Menj a **Tools & Settings → Bulk Actions → Scripts** menübe
3. Kattints a **+** gombra új script létrehozásához
4. Másold be a `pmax-asset-scheduler-v7.5.3.js` teljes tartalmát
5. Nevezd el a scriptet (pl. "PMax Asset Scheduler")

### 3. Konfiguráció beállítása

A script elején módosítsd a következő értékeket:

```javascript
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/XXXXXXX/edit';
const NOTIFICATION_EMAIL = 'te@email.com';
```

### 4. Ütemezés beállítása

1. A script szerkesztőben kattints a **Schedule** gombra
2. Állítsd be **óránkénti** futtatásra
3. Mentsd el az ütemezést

### 5. (Opcionális) YouTube API engedélyezése

A videó validációhoz:
1. A script szerkesztőben kattints a **Advanced APIs** gombra
2. Keresd meg a **YouTube Data API v3**-at
3. Kapcsold be

> **Megjegyzés**: A YouTube API nélkül is működik a script, de nem tudja előre ellenőrizni, hogy a videó elérhető-e.

---

## Konfiguráció

### Alapbeállítások

| Beállítás | Leírás | Alapérték |
|-----------|--------|-----------|
| `SPREADSHEET_URL` | A Google Sheet URL-je | *Kötelező* |
| `NOTIFICATION_EMAIL` | Email cím(ek) értesítéshez | *Kötelező* |
| `ADD_WINDOW_FROM_MIN` | ADD művelet kezdete (perc) | 0 (00:00) |
| `ADD_WINDOW_TO_MIN` | ADD művelet vége (perc) | 60 (01:00) |
| `REMOVE_WINDOW_FROM_MIN` | REMOVE művelet kezdete (perc) | 1380 (23:00) |
| `REMOVE_WINDOW_TO_MIN` | REMOVE művelet vége (perc) | 1440 (24:00) |
| `ALWAYS_SEND_PREVIEW` | Mindig küldjön preview email-t | `false` |

### Preview beállítások

Az `ALWAYS_SEND_PREVIEW` kapcsoló:
- **`false`** (alapérték): Csak akkor küld preview email-t, ha a sheet tartalma változott az előző futás óta
- **`true`**: **Mindig** küld preview email-t minden óránkénti futáskor
  - Hasznos rendszeres áttekintéshez (daily status report)
  - Megkerüli a sheet hash változás detektálást
  - Mindig friss állapot a jövőbeli műveletekről

### Post-Execution Preview

**Automatikus funkció** - Execution után újravalidál:

1. **Probléma**: Ha execution módosítja az állapotot, a korábbi preview elavult lehet
   - Példa: Ma REMOVE végrehajtódik → holnapi ADD még "már létezik" hibát mutatna (ROSSZ!)

2. **Megoldás**: Execution után automatikusan:
   - Újra lekéri az asset group állapotokat (friss adatok)
   - Újra validálja a jövőbeli sorokat
   - Preview email + Preview Results sheet a **valós** (post-execution) állapotot tükrözi

3. **Eredmény**: Pontosabb előrejelzés - a preview tükrözi, hogy az execution után mi a helyzet

### Időablakok

Ha az `Add Hour` vagy `Remove Hour` oszlopot üresen hagyod, a script alapértelmezett időablakot használ:
- **ADD műveletek**: alapértelmezetten **00:00-00:59** között hajtódnak végre (az adott dátumon)
- **REMOVE műveletek**: alapértelmezetten **23:00-23:59** között hajtódnak végre (az adott dátumon)

Ha kitöltöd az `Add Hour` / `Remove Hour` oszlopot (pl. `10`), a script az adott óra 00-59 perces ablakában hajt végre (pl. 10:00-10:59).

---

## Sheet Struktúra

### TextAssets

Szöveges asset-ek (headlines, long headlines, descriptions) ütemezése.

| Oszlop | Kötelező | Leírás |
|--------|----------|--------|
| Campaign Name | ✅ | A kampány pontos neve |
| Asset Group Name | ❌ | Asset group neve (üres = minden group) |
| Text Type | ✅ | `HEADLINE`, `LONG_HEADLINE`, vagy `DESCRIPTION` |
| Text | ✅ | A szöveg tartalma |
| Add Date | ❌ | Hozzáadás dátuma (YYYY-MM-DD) |
| Add Hour | ❌ | Hozzáadás órája (0-23). Ha üres: alapértelmezetten 00:00-00:59 között fut le |
| Remove Date | ❌ | Eltávolítás dátuma (YYYY-MM-DD) |
| Remove Hour | ❌ | Eltávolítás órája (0-23). Ha üres: alapértelmezetten 23:00-23:59 között fut le |

**Példa**:
| Campaign Name | Asset Group Name | Text Type | Text | Add Date | Add Hour | Remove Date | Remove Hour |
|---------------|------------------|-----------|------|----------|----------|-------------|-------------|
| PMax-Promo | Akciós termékek | HEADLINE | -50% Black Friday! | 2025-11-29 | 0 | 2025-12-02 | 23 |

### ImageAssets

Kép asset-ek ütemezése.

| Oszlop | Kötelező | Leírás |
|--------|----------|--------|
| Campaign Name | ✅ | A kampány pontos neve |
| Asset Group Name | ❌ | Asset group neve (üres = minden group) |
| Image Type | ✅ | `HORIZONTAL`, `SQUARE`, `VERTICAL 4:5`, vagy `VERTICAL 9:16` |
| Asset ID | ✅ | A kép Asset ID-ja (Google Ads-ból) |
| Add Date | ❌ | Hozzáadás dátuma |
| Add Hour | ❌ | Hozzáadás órája (0-23). Ha üres: alapértelmezetten 00:00-00:59 között fut le |
| Remove Date | ❌ | Eltávolítás dátuma |
| Remove Hour | ❌ | Eltávolítás órája (0-23). Ha üres: alapértelmezetten 23:00-23:59 között fut le |
| Thumbnail | 🤖 | *Automatikus* - Kattintható kép előnézet (a scripttől kapja az értékét) |

**Kép típusok és méretek**:
| Típus | Aspect Ratio | Példa méret |
|-------|--------------|-------------|
| HORIZONTAL | 1.91:1 | 1200×628 |
| SQUARE | 1:1 | 1200×1200 |
| VERTICAL 4:5 | 4:5 (0.8:1) | 960×1200 |
| VERTICAL 9:16 | 9:16 (0.5625:1) | 1080×1920 |

### VideoAssets

YouTube videó asset-ek ütemezése.

| Oszlop | Kötelező | Leírás |
|--------|----------|--------|
| Campaign Name | ✅ | A kampány pontos neve |
| Asset Group Name | ❌ | Asset group neve (üres = minden group) |
| Video Aspect (optional) | ❌ | `HORIZONTAL (16:9)`, `VERTICAL (9:16)`, vagy `SQUARE (1:1)` |
| Video URL or ID / Asset ID | ✅ | YouTube URL, Video ID, vagy Asset ID |
| Add Date | ❌ | Hozzáadás dátuma |
| Add Hour | ❌ | Hozzáadás órája (0-23). Ha üres: alapértelmezetten 00:00-00:59 között fut le |
| Remove Date | ❌ | Eltávolítás dátuma |
| Remove Hour | ❌ | Eltávolítás órája (0-23). Ha üres: alapértelmezetten 23:00-23:59 között fut le |
| Thumbnail | 🤖 | *Automatikus* - Kattintható videó előnézet (a scripttől kapja az értékét) |

**Videó azonosítás módjai** (mindhárom működik):
- YouTube URL: `https://www.youtube.com/watch?v=dQw4w9WgXcQ`
- YouTube Video ID: `dQw4w9WgXcQ`
- Google Ads Asset ID: `123456789`

> **Fontos - Video Aspect mező**: Ez az oszlop **csak informatív jellegű** - segít nyilvántartani, milyen típusú videókat adsz hozzá. A Google Ads API-ban a videóknak nincs altípusa (ellentétben a képekkel), minden videó `YOUTUBE_VIDEO` típusként kerül be. A script nem validálja és nem használja fel ezt a mezőt a műveletekhez.

---

## Használat

### 1. Töltsd ki a sheet-et

Add meg a kampány nevét, asset group-ot, asset tartalmat és dátumokat.

### 2. Várd meg a preview email-t

A script óránként fut és preview email-t küld a jövőbeli műveletekről.

### 3. Ellenőrizd a státuszokat

- **OK**: Minden rendben, végre fog hajtódni
- **WARNING**: Figyelmeztetés, de végre fog hajtódni
- **ERROR**: Hiba, nem fog végrehajtódni

### 4. Execution email

Amikor eljön az ütemezett idő, execution email-t kapsz a végrehajtott műveletekről.

---

## Linkek és Thumbnailok

### Bemeneti Sheet-ek (ImageAssets, VideoAssets)

A **Thumbnail oszlop** automatikusan feltöltődik minden futáskor — kattintható kép/videó előnézetek:

| Sheet | Mit mutat | Kattintva |
|-------|-----------|-----------|
| ImageAssets | A kép maga (thumbnail) | Megnyílik a kép teljes méretben |
| VideoAssets | YouTube videó thumbnail | Megnyílik a videó YouTube-on |

### Preview Results és Results Sheet-ek

A riport sheet-ekben szöveges linkek jelennek meg:

| Oszlop | Link |
|--------|------|
| Campaign | Nincs link — csak szöveg |
| Asset Group | Nincs link — csak szöveg |
| Asset ID (IMAGE) | 🖼️ Kattintható szöveges link → kép megnyílik böngészőben |
| Video | Nincs link — csak szöveg |

### Email Értesítések

Az emailben szöveges linkek szerepelnek (nem thumbnail képek — azok email kliensekben megbízhatatlanul töltenek be):

| Asset típus | Email tartalom |
|-------------|---------------|
| IMAGE | 🖼️ `assetId` — kattintható link a kép teljes méretű verziójára |
| VIDEO | Szöveg, link nélkül |
| TEXT | Szöveg — link nélkül |

### Hibakezelés

A link létrehozás **nem befolyásolja** a script működését:
- Ha egy link nem sikerül → hibalogolás, a többi link és a végrehajtás folytatódik
- Try-catch blokkok minden hyperlink beállításnál

---

## Státuszok és Hibaüzenetek

### Validációs státuszok

| Státusz | Jelentés |
|---------|----------|
| OK | Minden rendben |
| WARNING | Figyelmeztetés (pl. limit közel) |
| ERROR | Hiba — nem hajtódik végre |

### Végrehajtási státuszok

| Státusz | Jelentés |
|---------|----------|
| SUCCESS | Sikeres végrehajtás |
| SUCCESS [Verified ✓] | Sikeres és utólag ellenőrzött |
| ERROR | Sikertelen végrehajtás |

### Gyakori hibaüzenetek

| Hibaüzenet | Ok | Megoldás |
|------------|----|---------|
| Kampány nem található | Rossz kampánynév | Ellenőrizd a pontos kampánynevet |
| Asset ID nem található | Rossz Asset ID | Ellenőrizd az Asset ID-t a Google Ads-ban |
| Text asset már létezik | Duplikált szöveg | Töröld a duplikátumot a sheet-ből |
| Limit túllépés | Túl sok asset | Távolíts el asset-eket először |
| Video ID nem található asset-ként | REMOVE nem létező videóra | A videónak léteznie kell a fiókban |

---

## Limitek

> **Megjegyzés:** A limitek ellenőrzésekor a script csak a te által hozzáadott asset-eket számolja (`source = ADVERTISER`). A Google által automatikusan hozzáadott asset-ek nem számítanak bele a limitbe.

### Szöveges asset-ek (per asset group)

| Típus | Minimum | Maximum |
|-------|---------|---------|
| HEADLINE | 3 | 15 |
| LONG_HEADLINE | 1 | 5 |
| DESCRIPTION | 2 | 5 |

### Kép asset-ek (per asset group)

| Típus | Minimum | Maximum |
|-------|---------|---------|
| HORIZONTAL (1.91:1) | 1 | 20 |
| SQUARE (1:1) | 1 | 20 |
| VERTICAL 4:5 | 0 | 20 |
| VERTICAL 9:16 | 0 | 20 |

### Videó asset-ek (per asset group)

| Típus | Minimum | Maximum |
|-------|---------|---------|
| YOUTUBE_VIDEO | 0 | 5 |

---

## Gyakori Kérdések

### Miért nem törli a script a PAUSED asset group-ból az asset-et?

A script **csak ENABLED státuszú asset group-okkal** dolgozik. Ha egy asset group PAUSED állapotban van, a script figyelmen kívül hagyja. Ez szándékos viselkedés — szüneteltetett group-okra általában nem akarunk módosításokat végrehajtani.

### Miért kell óránként futtatni a scriptet?

A script időablakokban dolgozik (alapértelmezetten 00:00-00:59 és 23:00-23:59). Az óránkénti futtatás biztosítja, hogy a műveletek a megfelelő időablakban kerüljenek végrehajtásra.

### Lehet-e ugyanazon a napon ADD és REMOVE is?

Igen! Használd az `Add Hour` és `Remove Hour` oszlopokat különböző órák megadásához. Például:
- Add Hour: 10 (10:00-10:59 között hozzáadás)
- Remove Hour: 18 (18:00-18:59 között eltávolítás)

### Mi történik, ha a videó private vagy nem létezik?

Ha a YouTube API engedélyezve van, a script előre ellenőrzi a videó elérhetőségét:
- Private videó → ERROR
- Nem létező videó → ERROR
- Unlisted videó → OK (hirdetésben használható)

Ha a YouTube API nincs engedélyezve, WARNING-ot kapsz, de a művelet megpróbálódik.

### Asset ID hol található?

1. Google Ads → Tools & Settings → Asset Library
2. Keresd meg az asset-et
3. Az URL-ben vagy a részleteknél látható az Asset ID (számsor)

---

## Ismert Korlátozások

### Kampánylink nem érhető el

A Google Ads URL-ekben szereplő `ocid` paraméter egy belső, felhasználóhoz kötött azonosító — nem azonos a fiók customer ID-jával, és a scriptből nem kérhető le. Ennek következtében a script **nem generál kampánylinket** a riport sheet-ekben; a kampánynév egyszerű szövegként jelenik meg.

### Asset Group link nem érhető el

A Google Ads UI nem biztosít közvetlen deep link URL-t asset group-okhoz. Az asset group neve egyszerű szövegként jelenik meg.

### Videó link Asset ID formátumnál nem generálódik

Ha a VideoAssets fülön a videó **Asset ID formátumban** van megadva (pl. `312097984366`), a script nem tud YouTube linket generálni — ehhez az API-ból kellene lekérnie a YouTube videó ID-t az asset ID alapján. Ez jövőbeli fejlesztési lehetőség.

YouTube URL vagy video ID formátumnál (pl. `dQw4w9WgXcQ`) a link generálódik.

---

## Verzió

**Jelenlegi verzió**: v7.5.3

Lásd a [CHANGELOG_Version2.md](./CHANGELOG_Version2.md) fájlt a részletes változásokért.

---

## Készítők

© 2025 Klára Bognár – [Impresszio Online Marketing](https://impresszio.hu)

Fejlesztve Claude Code segítségével, Google Ads Script Sensei © Nils Rooijmans tanácsaival.
