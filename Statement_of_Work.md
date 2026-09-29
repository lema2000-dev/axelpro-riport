# AxelPro napi riport automatizálása

## Cél

Windows alatt futó, egyszerű grafikus Python-alkalmazás, amely egy AxelPro-ból exportált `Mozgások.csv` alapján kitölti a napi értékesítési riportot. A fejlesztés Ubuntu és VS Code alatt történik, a Windowsos működést külön teszteljük. A mintafájl: `Napi értékesítési riport 2026.xlsx`.

## Felhasználói folyamat

1. Az első indításkor a felhasználó kiválaszt egy mentési mappát, és létrehozza az adott év üres riportját. Ha ez nem januárban történik, a számítógép aktuális hónapjáig **januártól kezdve az összes havi munkalapot** létrehozzuk; a korábbi hónapok adatbeviteli cellái üresek maradnak. Az `Összesítő` a létrehozott hónapok oszlopait tartalmazza; az `Euro` lapot is előkészítjük.
2. A program megjegyzi a mentési mappát, és következő indításkor ott keresi a riportot. Ha a fájl hiányzik vagy több megfelelő fájl van, egyértelmű választást kér.
3. A felhasználó kiválasztja a CSV-exportot. A program a **CSV tételeinek dátuma szerint** választja ki a célhónapot, célnapot és cél évet, előnézetben megmutatja az írandó összegeket és az esetleges kihagyott sorokat, majd menti a frissített Excel-fájlt. Egy CSV több napot vagy hónapot is tartalmazhat.
4. Az új hónap első, még nem feldolgozott tételénél létrehozza az adott havi munkalapot, megtartva a címkéket, formázást és képleteket, és hozzáadja a hozzá tartozó havi oszlopot az `Összesítő` laphoz. A már létező hónapot vagy összesítő oszlopot nem duplikálja. Ha egy korábbi, hiányzó hónapot kell pótolni, az azt megelőző év elejétől a célhónapig szükséges üres lapokat és összesítő oszlopokat is létrehozza. Ha új év első tétele érkezik, **külön éves Excel-fájlt** hoz létre januári lappal, `Euro` lappal és az `Összesítő` kizárólag januári havi oszlopával; a későbbi hónapok az első oda tartozó tételnél kerülnek bele.
5. A felületen szerkeszthető, hogy az `Új alkatrész` mellett mely további `Csoport` értékek számítsanak új alkatrésznek.
6. A felületen külön **Alaphelyzetbe állítás** művelet törli a program saját mentett állapotát (beállítások, megjegyzett mappa, felvett csoportok, árfolyamgyorsítótár és feldolgozott bizonylatok). A már létrehozott Excel-riportokat nem módosítja; a művelet előtt egyértelműen jelzi, hogy régi CSV újraimportja ezután duplázhatja az Excelben már szereplő értékeket.

## Bemenet és összesítési szabályok

A minta CSV UTF–16 kódolású, tabulátorral tagolt fájl. Az összegekhez mindig a **`Ne. érték` (nettó érték)** mezőt használjuk, tizedes vesszővel. A HUF és EUR értékeket külön tartjuk nyilván; a havi munkalapon egy naphoz forint- és euróoszlop tartozik. Az érintett napi cellákat a hónap munkalapján, a dátumot tartalmazó 2. sor alapján azonosítjuk, nem rögzített oszlopbetűkkel. A `Csoport` nevek összehasonlítása kis- és nagybetűtől független; a felhasználó által felvett további csoportoknál is.

| Cél a havi munkalapon | CSV-feltétel | Művelet |
| --- | --- | --- |
| 8. sor – Bontott alkatrész sz. | `Ügylet típus = Számla`, `Csoport = Bontott alkatrész` | Napi nettó értékek összege. |
| 9. sor – Új alkatrész sz. | `Ügylet típus = Számla`, `Csoport = Új alkatrész` vagy felvett további csoport | Napi nettó értékek összege. |
| 10. sor – Munkadíj sz. | `Ügylet típus = Számla`, `Csoport = Munkadíj` | Napi nettó értékek összege. |
| 23. sor – Alkatrész bontott műhely | Olyan számla bontottalkatrész-tételei, amelyen munkadíjtétel is van | Napi nettó értékek összege; egy bizonylatot csak egyszer veszünk figyelembe. |
| 24. sor – Alkatrész új műhely | Olyan számla újalkatrész-tételei, amelyen munkadíjtétel is van | Napi nettó értékek összege; az újként felvett csoportok is ide tartoznak. |
| 27. sor – Kiadás alkatrész sz. | `Ügylet típus = Beszerzés` | Napi nettó értékek összege. |
| 32. sor – Posta bevétel | `Csoport = Szállítási költség` | Napi nettó értékek összege. |
| 43. sor – Trailer bérbeadás | `Csoport = Trailer bérbeadás` | Napi nettó értékek összege; ez a csoport a jelenlegi mintában még nincs. |
| Három új, külön havi sor: `SANY`, `DIAG`, `KJ` | `Ügylet típus = Számla`, és `Cikkszám = SANY`, `DIAG`, illetve `KJ` | **Cikkszámonként külön havi** nettó összeg, amelyhez minden új, még fel nem dolgozott számlatétel egyszer számít hozzá. |

A jelenlegi havi lap 106. sorában még egyetlen összevont `Segédanyag, diag., klíma` cella van; ezt a sablonban három külön sor váltja fel. A sorok beszúrása miatt a 106. sor alatti képleteket, formázásokat és az esetleges összesítő hivatkozásokat át kell vezetni. Az `Összesítő` is három külön havi eredményt és hozzájuk tartozó éves összeget mutasson. A többi Excel-képletet, a heti és havi zárásokat a riport saját képletei végzik. A feldolgozó csak a megállapodott bemeneti cellákat módosítja. A cikkszámok összehasonlítását is kis- és nagybetűtől függetlenül tervezzük.

## Dátum, bizonylatok, árfolyam

- A cél napját és évét minden tételnél a CSV `Árumozgás időpontja` mezője határozza meg; a gép dátuma kizárólag az új riport létrehozásakor szükséges kezdeti hónaphoz.
- A program **minden feldolgozott bizonylatszámot az évszámmal együtt** tartósan megőriz, például `2026 / TIM1-SZ-1790990` alakban jelenít meg. Futás közben memóriában kezeli, újraindítás után pedig helyi állományból tölti vissza. A már ismert év–bizonylatszám páros teljes tételsora kimarad az új importból, így a havi `SANY`, `DIAG`, `KJ` összegek sem duplázódnak. Üres bizonylatszám a tényleges exportban nem fordulhat elő; ha mégis szerepel, a program hibát jelez és nem könyveli el. A feldolgozási állapotot a munkafüzettel összehangoltan mentjük, hogy félbeszakadt mentésnél se vesszen el vagy duplázódjon tétel.
- Egy CSV több évet is tartalmazhat. Ilyenkor minden új bizonylat a **saját évének Excel-fájljába**, azon belül a CSV dátumához tartozó hónapba és napba kerül. Az import eredményét évenként és célfájlonként külön mutatjuk meg.
- Az EUR/HUF árfolyamot a **Magyar Nemzeti Bank hivatalos középárfolyamának webszolgáltatásából** kérjük le a tétel dátumához, és a megfelelő éves munkafüzet `Euro` lapjára rögzítjük. Ha az adott napra még nem érhető el friss MNB-árfolyam vagy a lekérés sikertelen, **kézi beírási lehetőséget ajánl fel**; jóváhagyott árfolyam nélkül nem könyvel el érintett eurótételeket. A használt értéket és forrást naplózza.
- A forintban közölt végösszegeket **egész forintra** kerekítjük; az EUR-értékeket euróban tartjuk, és az EUR-ból számított forintösszegek is egész forintra kerekítve jelennek meg. Az összesítést `Decimal` értékeken végezzük, és az összegzés után kerekítünk, hogy a tételek egyenkénti kerekítése ne halmozzon hibát. A végleges képletekben egységesen rögzítjük a kerekítési módot.

## Fejlesztési munkamenet

### 1. Projekt és környezet

- VS Code projektmappa, Python `.venv`, kiválasztott interpreter.
- Függőségek: `openpyxl`, `pytest`; a felülethez kezdetben `tkinter`, a CSV-hez és hálózathoz Python standard könyvtári modulok.
- Projektfájlok: `src/` az alkalmazáshoz, `tests/` a fontos feldolgozási esetekhez, `samples/` kizárólag személyes adatoktól megtisztított példákhoz. A valódi ügyféladat és a helyi beállítás ne kerüljön nyilvános Git-re.

### 2. A minta és a szabályok pontosítása

- Feltérképezzük a hónaplapok eltérő hosszát, a napi oszlopokat, a három külön havi sor helyét, az összesítő kapcsolatait és az `Euro` lap képlethivatkozásait.
- Ellenőrizzük a kézi MNB-árfolyam megadásának felületét, az egész forintra kerekítő képleteket, az év–bizonylatszám kulcsot és a több éves importot; a `SANY`, `DIAG`, `KJ` cikkszámok kizárólag `Számla` ügylettel, a CSV szerinti dátum és minden `Beszerzés` csoport már rögzített szabály.
- Kézzel kiszámolunk néhány reprezentatív számlát és egy teljes mintanapot az ellenőrzéshez.

### 3. Feldolgozó mag

- CSV beolvasása és validálása; pénzösszegek kezelése `Decimal` típussal.
- Tételek csoportosítása CSV-dátum, év, pénznem, kisbetűfüggetlen csoport, cikkszám és számlaszám szerint; a műhelyhez kapcsolódó számlák kiválasztása. A `Beszerzés` minden csoportját összegezzük; a `Trailer bérbeadás` csoportnál nincs `Számla` szűrés.
- Napi és havi összegek számítása, valamint az importok közötti ismétlődés kiszűrése. A feldolgozás eredménye írás előtt áttekinthető legyen.

### 4. Excel és tartós állapot

- A gép aktuális évének januárjától az aktuális hónapig terjedő, szükség esetén üres munkalapok létrehozása, a három új havi sor kialakítása és az `Összesítő` kezdeti havi oszlopainak összekötése. A következő hónap és összesítő oszlop létrehozása az első oda tartozó CSV-tételnél; új évnél új éves munkafüzet januári kezdettel, régebbi év pótlásakor az érintett év saját fájljának bővítése. A dátumok, formázás és minden érintett képlet ellenőrzése.
- Csak a célcellák módosítása; a meglévő munkafüzet biztonsági másolatának készítése mentés előtt.
- Beállítások és az **összes eddig feldolgozott év–bizonylatszám pár** tartós helyi mentése; következetes működés újraindítás, átfedő export, több éves CSV és évváltás esetén. Külön reset művelet a teljes alkalmazásállapotra, az Excel-fájlok megőrzésével.

### 5. Grafikus felület és árfolyam

- Mentési mappa, riport létrehozása, CSV kiválasztása, további újalkatrész-csoportok szerkesztése és saját állapot alaphelyzetbe állítása.
- Feldolgozási előnézet, érthető hibaüzenetek, sikeres mentés visszajelzése.
- Árfolyam API és kézi tartalék megadás; a felhasznált adat rögzítése.

### 6. Ellenőrzés és átadás

- Ellenőrző esetek: több tételes számla, több munkadíj ugyanazon számlán, HUF/EUR, egész forintos kerekítés, kis- és nagybetűs csoportnév, három elkülönített számlás cikkszám, átfedő és több évet tartalmazó export, újraindítás, januárnál későbbi első indítás, hónap- és évváltás, hiányzó árfolyam és kézi bevitel, reset utáni újraimport, hiányzó vagy hibás munkafüzet.
- A kézzel számolt mintaösszegek és az Excel-képletek eredményének összevetése.
- Windows VM-en GUI- és Excel-mentési próba; ezután Windowsos futtatható csomag és rövid használati útmutató.

## Megvalósításkor ellenőrzendő részletek

- A mintafájl EUR→HUF képletei jelenleg nem feltétlenül kerekítenek. A három új havi sor és az összesítő mellett ezeket is az egész forintos szabályhoz kell igazítani.
- A kézi árfolyamot csak a hiányzó napi értékhez kérjük; ha az adott naphoz egyszer már rögzítettünk árfolyamot, az ismételt import ugyanazt használja.
- Az év–bizonylatszám párok, az Excel-módosítások és az árfolyamok összehangolt mentését ellenőrizni kell akkor is, ha egy CSV több éves fájlt módosít.
- A reset után meglévő Excel-riportba történő régi CSV-import kétszeres összegzést okozhat. A felületen ezt a művelet előtt jelezzük; biztonságos újrakezdéshez új, üres riportot kell létrehozni vagy a meglévőt egyeztetni.
