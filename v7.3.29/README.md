# PMax Asset Scheduler

**Verzió:** v7.3.29
**Szerző:** Klára Bognár (Impresszió Online Marketing)
**Weboldal:** [impresszio.hu](https://impresszio.hu)

---

## Leírás

Automatizált Google Ads script Performance Max kampányok asset kezeléséhez. Időzített TEXT és IMAGE asset hozzáadás/törlés asset group-ok szerint.

**Funkciók:**
- ⏰ Időzített asset hozzáadás és törlés
- 📊 Előnézeti validáció (30 napos horizont)
- ✅ Végrehajtás utáni ellenőrzés
- 📧 Email értesítések (preview és execution)
- 📋 Google Sheets jelentések

✅ Szöveg asset-ek (headline, long headline, description) időzített hozzáadása/törlése
✅ Kép asset-ek időzített hozzáadása/törlése (Asset ID alapján)
✅ Rugalmas ütemezés: bármely órára beütemezhetők műveletek (pl. 10:00-kor ADD, 19:00-kor REMOVE)
✅ Asset Group szintű vezérlés: konkrét group
✅ Ha csak kampánynevet tartalmaz egy sor, minden aktív, nem feed-only elemcsoportra hajtja végre a műveleteket
✅ Automatikus validáció: ellenőrzi a Google limiteket (csak hirdető által feltöltött asset-eket számolva), duplikációkat, hibás beállításokat
✅ Preview email: jövőbeli műveletek előnézete (30 napra előre)
✅ Execution email: végrehajtott műveletek összesítése
✅ Safety: üzleti logika előnyben (lejárt promóciós tartalom törlése fontosabb mint a min limitek)

❌ Nem csinál automatikus képfeltöltést - az Asset ID-t előre fel kell töltened a Google Ads-be
❌ Nem tudja előre látni a jövőbeli kampány változásokat
❌ Nem tudja több művelet jövőbeli hatását összegzetten ellenőrizni egy elemcsoportra
❌ Nem kezeli a Google moderációt - a script hozzáadja az asset-et, de a Google csak ezután ellenőrzi

---

## Telepítés

### 1. Google Sheet létrehozása

**Template másolása:**
[📄 Google Sheet Template](https://docs.google.com/spreadsheets/d/1HHWrSD8pCP87u63bDfFBDyqKIFwUh3tX-qpfXmME_hs/copy)

**Szükséges fülek a template-ben:**
- `TextAssets` - szöveges asset-ek ütemezése
- `ImageAssets` - kép asset-ek ütemezése

**Megjegyzés:** A `Preview Results` és `Results` füleket a script automatikusan létrehozza az első futás során. A fülek sorrendje utána szabadon módosítható.

### 2. Script telepítése Google Ads-be

1. Jelentkezz be a Google Ads fiókodba
2. Navigálj a **Tools > Bulk Actions > Scripts** menüpontra
3. Kattints a **+ NEW SCRIPT** gombra
4. Másold be a `pmax-asset-scheduler-v7.3.29.js` tartalmát
5. Mentsd el a scriptet

### 3. Konfiguráció

Állítsd be a script elején található konfigurációs változókat:

```javascript
// Kötelező: Google Sheet URL
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/';

// Kötelező: Értesítési e-mail(ek), vesszővel elválasztva
const NOTIFICATION_EMAIL = 'your-email@example.com';
```

### 4. Ütemezés beállítása

A scriptnek **óránként kell futnia** az időzített műveletek végrehajtásához:

---

## Időablakok

Az alapértelmezett végrehajtási időablakok (a konfig szekcióban módosíthatók):

- **ADD műveletek:** 00:00-00:59 
- **REMOVE műveletek:** 23:00-23:59 

---

## Használat

1. Töltsd ki a Google Sheet-et az ütemezendő asset-ekkel
2. A script óránként fut és:
   - **Preview mode:** ellenőrzi a jövőbeli műveleteket (30 nap), ha a sheetbe ellenőrizhető műveletek kerültek a script előző futása óta.
   - **Execution mode:** végrehajtja az időablakban lévő műveleteket
3. Email értesítéseket kapsz az eredményekről
4. A `Preview Results` és `Results` füleken láthatod a részleteket

---

## Dokumentáció

- **Changelog:** [CHANGELOG_Version2.md](./CHANGELOG_Version2.md) - Teljes verziótörténet
- **Script forráskód:** [pmax-asset-scheduler-v7.3.29.js](./pmax-asset-scheduler-v7.3.29.js)

---

⚠️ FONTOS: Ha asset cserét szeretnél, és az asset típusból a maximális elemszám már megvan az elemcsoportban (REMOVE + ADD ugyanolyan típusból), 
akkor legalább 1 órával a hozzáadás elé időzítsd a törlés műveletet! 

Jó példa: REMOVE 9:00, ADD 10:00
Rossz példa: REMOVE 10:00, ADD 10:00 (ugyanaz az óra)

A magyrázat, hogy a végrehajtás órájában is lefut egy elővalidáció, ami kizárja a végrehajtásból a max. elemszámok fölötti műveleteket.


## Kapcsolat

**Impresszió Online Marketing**
🌐 [impresszio.hu](https://impresszio.hu)

