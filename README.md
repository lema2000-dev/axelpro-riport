# AxelPro napi riport automatizálása

## Cél

Windows alatt futó, egyszerű grafikus Python-alkalmazás, amely egy AxelPro-ból exportált `Mozgások.csv` alapján kitölti a napi értékesítési riportot. A fejlesztés Ubuntu és VS Code alatt történik, a Windowsos működést külön teszteljük. A végleges sablon forrása: `Másolat - Napi értékesítési riport 2026(1).xlsx`, annak **Szeptember** munkalapján az első hét napi és heti záró része, valamint a havi záró. A későbbi heti blokkok még korábbi változatot tartalmaznak, ezért nem másolhatók változtatás nélkül.

## Felhasználói folyamat

1. Az első indításkor a felhasználó kiválaszt egy mentési mappát, és létrehozza az adott év üres riportját. Ha ez nem januárban történik, a számítógép aktuális hónapjáig **januártól kezdve az összes havi munkalapot** létrehozzuk; a korábbi hónapok adatbeviteli cellái üresek maradnak. Az `Összesítő` a létrehozott hónapok oszlopait tartalmazza; az `Euro` lapot is előkészítjük.
2. A program megjegyzi a mentési mappát, és következő indításkor ott keresi a riportot. Ha a fájl hiányzik vagy több megfelelő fájl van, egyértelmű választást kér.
3. A felhasználó kiválasztja a CSV-exportot. A program a **CSV tételeinek dátuma szerint** választja ki a célhónapot, célnapot és cél évet, előnézetben megmutatja az írandó összegeket és az esetleges kihagyott sorokat, majd menti a frissített Excel-fájlt. Egy CSV több napot vagy hónapot is tartalmazhat.
4. Az új hónap első, még nem feldolgozott tételénél létrehozza az adott havi munkalapot, megtartva a címkéket, formázást és képleteket, és hozzáadja a hozzá tartozó havi oszlopot az `Összesítő` laphoz. A már létező hónapot vagy összesítő oszlopot nem duplikálja. Ha egy korábbi, hiányzó hónapot kell pótolni, az azt megelőző év elejétől a célhónapig szükséges üres lapokat és összesítő oszlopokat is létrehozza. Ha új év első tétele érkezik, **külön éves Excel-fájlt** hoz létre januári lappal, `Euro` lappal és az `Összesítő` kizárólag januári havi oszlopával; a későbbi hónapok az első oda tartozó tételnél kerülnek bele.
5. A felületen szerkeszthető, hogy az `Új alkatrész` mellett mely további `Csoport` értékek számítsanak új alkatrésznek.
6. A felületen külön **Alaphelyzetbe állítás** művelet törli a program saját mentett állapotát (beállítások, megjegyzett mappa, felvett csoportok, árfolyamgyorsítótár és feldolgozott bizonylatok). A már létrehozott Excel-riportokat nem módosítja; a művelet előtt egyértelműen jelzi, hogy régi CSV újraimportja ezután duplázhatja az Excelben már szereplő értékeket.

## Bemenet és összesítési szabályok

A minta CSV UTF–16 kódolású, tabulátorral tagolt fájl. Az összegekhez mindig a **`Ne. érték` (nettó érték)** mezőt használjuk, tizedes vesszővel. A pénznemet a `Pénznem` mezőből olvassuk ki. Az EUR-tételeket a CSV szerinti nap EUR/HUF árfolyamával forintra váltjuk, majd a HUF-tételekkel együtt összegezzük. **Minden automatikus adatbevitel forintban történik.** A sablon kézzel kitöltött részein megmaradó eurócellák nem a CSV-import célcellái. Az érintett napi cellákat a hónap munkalapján, a dátumot tartalmazó 2. sor alapján azonosítjuk, nem rögzített oszlopbetűkkel. A `Csoport` nevek összehasonlítása kis- és nagybetűtől független; a felhasználó által felvett további csoportoknál is.

| Cél a havi munkalapon | CSV-feltétel | Művelet |
| --- | --- | --- |
| 8. sor – Bontott alkatrész sz. | `Ügylet típus = Számla`, `Csoport = Bontott alkatrész` | Napi nettó értékek összege. |
| 9. sor – Új alkatrész sz. | `Ügylet típus = Számla`, `Csoport = Új alkatrész` vagy felvett további csoport | Napi nettó értékek összege. |
| 10. sor – Munkadíj sz. | `Ügylet típus = Számla`, `Csoport = Munkadíj` | Napi nettó értékek összege. |
| 23. sor – Alkatrész bontott műhely | Olyan számla bontottalkatrész-tételei, amelyen munkadíjtétel is van | Napi nettó értékek összege; egy bizonylatot csak egyszer veszünk figyelembe. |
| 24. sor – Alkatrész új műhely | Olyan számla újalkatrész-tételei, amelyen munkadíjtétel is van | Napi nettó értékek összege; az újként felvett csoportok is ide tartoznak. |
| 27. sor – Kiadás alkatrész sz. | `Ügylet típus = Beszerzés` | Napi nettó értékek összege. |
| 32. sor – Posta bevétel | `Csoport = Szállítási költség` | Napi nettó értékek összege. |
| 40. sor – Trailer bérbeadás | `Csoport = Trailer` | Napi nettó értékek összege; a mintában szereplő EUR-tételt is átváltjuk. |
| B82 – Segédanyag | `Ügylet típus = Számla`, `Cikkszám = SANY` | Havi nettó összeg forintban. |
| B83 – Diagnosztika | `Ügylet típus = Számla`, `Cikkszám = DIAG` | Havi nettó összeg forintban. |
| B84 – Klímajavítás | `Ügylet típus = Számla`, `Cikkszám = KJ` | Havi nettó összeg forintban. |

A három havi összeghez minden új, még fel nem dolgozott számlatétel egyszer számít hozzá. A cikkszámok összehasonlítása is kis- és nagybetűtől független.

## Végleges Excel-sablon és összesítő

- Az első heti blokk napi oszlopainak kezdőcellái: `B`, `D`, `F`, `H`, `J`, `L`; a heti záró `N:O`. Az automatikus napi célcellák a végleges sablonban összevont forintcellák. A többi heti blokkot ebből generáljuk, a dátumokat és képlethivatkozásokat a célhéthez igazítva.
- A havi záró a 43–85. sorban található; értékei a `B:C` összevont cellákban vannak. A havi képletek a létrehozott heti blokkokra és a saját hónap árfolyamcellájára hivatkoznak.
- Ahol a végleges mintában képlet van, ott annak megfelelő képletet generálunk. Ahol nincs képlet, és a cella nem szerepel az automatikus kitöltési szabályok között, az **kézi adatbevitelre marad**. Új riportba nem másoljuk át a minta kézzel bevitt üzleti értékeit; importáláskor a felhasználó meglévő kézi adatait megőrizzük.
- Az `Összesítő` a havi záró sorfeliratait és sorsorrendjét követi. A hónapok **egymás melletti oszlopokban** jelennek meg; minden érték képlettel hivatkozik a megfelelő havi záró cellára, például `='Szeptember'!B46`. Új hónap munkalapjával együtt létrejön annak összesítő oszlopa is. Nem készítünk ettől eltérő új összesítési logikát.
- Az egyszerűsített `Euro` lapon minden hónap neve mellett **egyetlen EUR/HUF árfolyamcella** van. A havi záró kézi EUR-adatait használó képletek erre hivatkoznak; a régi, napokra bontott Euro-sablont nem használjuk.

## Dátum, bizonylatok, árfolyam

- A cél napját és évét minden tételnél a CSV `Árumozgás időpontja` mezője határozza meg; a gép dátuma az új riport kezdeti hónapjához és az aktuális havi záró árfolyamának napjához szükséges.
- A program **minden feldolgozott bizonylatszámot az évszámmal együtt** tartósan megőriz, például `2026 / TIM1-SZ-1790990` alakban jelenít meg. Futás közben memóriában kezeli, újraindítás után pedig helyi állományból tölti vissza. A már ismert év–bizonylatszám páros teljes tételsora kimarad az új importból, így a havi `SANY`, `DIAG`, `KJ` összegek sem duplázódnak. Üres bizonylatszám a tényleges exportban nem fordulhat elő; ha mégis szerepel, a program hibát jelez és nem könyveli el. A feldolgozási állapotot a munkafüzettel összehangoltan mentjük, hogy félbeszakadt mentésnél se vesszen el vagy duplázódjon tétel.
- Egy CSV több évet is tartalmazhat. Ilyenkor minden új bizonylat a **saját évének Excel-fájljába**, azon belül a CSV dátumához tartozó hónapba és napba kerül. Az import eredményét évenként és célfájlonként külön mutatjuk meg.
- Az EUR/HUF árfolyamot a **Magyar Nemzeti Bank hivatalos középárfolyamának webszolgáltatásából** kérjük le. Ha a kért napra nem érhető el árfolyam vagy a lekérés sikertelen, **kézi beírási lehetőséget ajánlunk fel**, az alábbi múltbeli pótlási szabály figyelembevételével. A használt értéket, dátumot és forrást tartósan rögzítjük.
- **CSV-tételek átváltása:** a tétel CSV-dátumához tartozó napi árfolyamot használjuk. Az ugyanahhoz a naphoz már rögzített árfolyam ismételt importnál is érvényes; ezek a napi adatok az alkalmazás saját állapotában maradnak. Árfolyam nélkül nem könyveljük el az érintett EUR-tételeket.
- **Hiányzó múltbeli napi árfolyam:** ha a CSV-tétel napja korábbi a gép aktuális napjánál, és nincs hozzá rögzített vagy lekérhető napi árfolyam, az adott év **azonos hónapjának időben legközelebbi elérhető napi árfolyamát** használjuk. Ez lehet a célnap előtti vagy utáni nap; azonos távolságnál a korábbi napot választjuk. Más hónapból nem pótolunk. Ha ugyanabban a hónapban nincs elérhető napi adat, kézi bevitelt kérünk. A pótlást a célnaphoz tartósan eltároljuk az eredeti árfolyamdátummal és forrással együtt, így később ugyanazt használjuk. A pótlási kereséshez tényleges napi vagy kézzel rögzített adatot használunk, másik napra már átmásolt pótlást nem tekintünk új napi forrásnak. Az aktuális nap hiányzó friss árfolyamára továbbra is kézi bevitelt ajánlunk fel.
- **Havi záró árfolyama:** az aktuális hónap `Euro`-celláját a program használatakor az aktuális napi MNB-árfolyammal vagy kézi értékkel frissítjük. Ez a havi záró kézzel bevitt EUR-összegeinek átváltásához kell. A hónap lezárulta után az utolsó, abban a hónapban rögzített árfolyam megmarad, akkor is, ha kézzel adták meg. Korábbi hónap CSV-jének későbbi importja nem írja át ezt a lezárt havi árfolyamot.
- A forintban beírt végösszegeket **egész forintra** kerekítjük. `Decimal` értékekkel számolunk: előbb a szükséges napi árfolyamokkal átváltunk és összegezünk, utána kerekítünk. A fél forintos értékeket nullától távolodva kerekítjük (`ROUND_HALF_UP`): `12,5 → 13`, `−12,5 → −13`. A sablon jóváhagyott üzleti képleteit megőrizzük; esetleges kerekítési módosításukat külön egyeztetjük.

- **Ismételt napi vagy havi import:** kizárólag az új bizonylatok adott célcellához tartozó összegeit számítjuk ki, ezt a növekményt kerekítjük, majd hozzáadjuk az Excelben már szereplő értékhez. A feldolgozott tételek összegeit és a kerekítési maradékokat nem tároljuk tartósan az alkalmazás saját állapotában. A részletekben végzett import eredménye ezért eltérhet az egyszeri teljes importétól; ezt elfogadott működésként kezeljük. A bizonylatazonosítókat és napi árfolyamokat továbbra is megőrizzük.
- **Kézi árfolyambevitel megszakítása:** a teljes import mentés nélkül leáll. Sem a riportok, sem a feldolgozott bizonylatok nyilvántartása nem változik; így később az egész fájl újrapróbálható.

## Fejlesztési munkamenet

### 1. Projekt és környezet – kész

- VS Code projektmappa, Python `.venv`, kiválasztott interpreter.
- Függőségek: `openpyxl`, `pytest`; a felülethez kezdetben `tkinter`, a CSV-hez és hálózathoz Python standard könyvtári modulok.
- Projektfájlok: `src/` az alkalmazáshoz, `tests/` a fontos feldolgozási esetekhez, `samples/` kizárólag személyes adatoktól megtisztított példákhoz. A valódi ügyféladat és a helyi beállítás ne kerüljön nyilvános Git-re.

### 2. A minta és a szabályok pontosítása – lezárva

- **Tisztázva:** a végleges első heti blokk, heti és havi záró; az automatikus napi és havi célcellák; a kézi mezők megőrzése; az összesítő havi képlethivatkozásai; az egyszerűsített Euro lap és a két külön árfolyamfelhasználás.
- **Tisztázva:** CSV szerinti dátum, tartós év–bizonylatszám kulcs, több éves import, minden beszerzési csoport, trailer számlaszűrés nélkül, a három külön számlás cikkszám és a reset hatóköre.
- Az eredeti mintanap és reprezentatív bizonylatai összesítését ellenőriztük. Elkészültek a fiktív többnapos, többhónapos és többéves CSV-k; várt nyers összegeik alább szerepelnek. A tényleges program és Excel ellenőrzése a fejlesztés során ezekre épül.
- Végleges: egész forintos, nullától távolodó félérték-kerekítés; importonként kerekített növekmény, összegek tartós tárolása nélkül; megszakított kézi árfolyambevitelkor teljes import mentés nélküli leállítása.
- Következő szakasz: **3. Feldolgozó mag**, fájlonként haladva.

### 3. Feldolgozó mag

- CSV beolvasása és validálása; pénzösszegek kezelése `Decimal` típussal.
- Tételek csoportosítása CSV-dátum, év, pénznem, kisbetűfüggetlen csoport, cikkszám és számlaszám szerint; a műhelyhez kapcsolódó számlák kiválasztása. A `Beszerzés` minden csoportját összegezzük; a `Trailer` csoportnál nincs `Számla` szűrés.
- Napi és havi összegek számítása, valamint az importok közötti ismétlődés kiszűrése. A feldolgozás eredménye írás előtt áttekinthető legyen.

### 4. Excel és tartós állapot

- A gép aktuális évének januárjától az aktuális hónapig terjedő, szükség esetén üres munkalapok létrehozása, a végleges havi sablon alkalmazása és az `Összesítő` kezdeti havi oszlopainak összekötése. A következő hónap és összesítő oszlop létrehozása az első oda tartozó CSV-tételnél; új évnél új éves munkafüzet januári kezdettel, régebbi év pótlásakor az érintett év saját fájljának bővítése. A dátumok, formázás és minden érintett képlet ellenőrzése.
- Heti blokkok és havi záró képleteinek generálása a végleges mintából; az összesítő havi hivatkozásainak és a hónaponként egycellás Euro lapnak a létrehozása.
- Csak a célcellák módosítása; a meglévő munkafüzet biztonsági másolatának készítése mentés előtt.
- Beállítások és az **összes eddig feldolgozott év–bizonylatszám pár** tartós helyi mentése; következetes működés újraindítás, átfedő export, több éves CSV és évváltás esetén. Külön reset művelet a teljes alkalmazásállapotra, az Excel-fájlok megőrzésével.

### 5. Grafikus felület és árfolyam

- Mentési mappa, riport létrehozása, CSV kiválasztása, további újalkatrész-csoportok szerkesztése és saját állapot alaphelyzetbe állítása.
- Feldolgozási előnézet, érthető hibaüzenetek, sikeres mentés visszajelzése.
- MNB-árfolyam lekérése és kézi tartalék megadás; CSV-napi árfolyamok tárolása, aktuális havi Euro-cella frissítése és lezárt havi árfolyamok megőrzése.

### 6. Ellenőrzés és átadás

- Ellenőrző esetek: több tételes számla, több munkadíj ugyanazon számlán, HUF/EUR, egész forintos kerekítés, kis- és nagybetűs csoportnév, három elkülönített számlás cikkszám, átfedő és több évet tartalmazó export, újraindítás, januárnál későbbi első indítás, hónap- és évváltás, hiányzó árfolyam és kézi bevitel, reset utáni újraimport, hiányzó vagy hibás munkafüzet.
- A kézzel számolt mintaösszegek és az Excel-képletek eredményének összevetése.
- Windows VM-en GUI- és Excel-mentési próba; ezután Windowsos futtatható csomag és rövid használati útmutató.

## Megvalósításkor ellenőrzendő részletek

- Az új havi lapok heti blokkjainak száma és képlethivatkozásai igazodjanak a célhónap naptárához. A hónaphatáron átnyúló heti blokkba csak a saját hónap napjait generáljuk. A következő havi lapon ugyanennek a hétnek a folytatása azonos naptári hétszámot kap; a heti záró kizárólag a saját havi blokk napjait összegzi.
- A napi átváltási adatok és a havi záró egyetlen árfolyama külön szerepet töltenek be; ismételt import és korábbi hónap pótlása során sem keveredhetnek.
- Ellenőrizzük az összesítő képleteit, a kézi EUR-mezők havi átváltását, a lezárt havi árfolyam megőrzését és az automatikus cellákon kívüli kézi adatok változatlanságát.
- Az év–bizonylatszám párok, az Excel-módosítások és az árfolyamok összehangolt mentését ellenőrizni kell akkor is, ha egy CSV több éves fájlt módosít.
- A reset után meglévő Excel-riportba történő régi CSV-import kétszeres összegzést okozhat. A felületen ezt a művelet előtt jelezzük; biztonságos újrakezdéshez új, üres riportot kell létrehozni vagy a meglévőt egyeztetni.

## Fiktív CSV-ellenőrző minták

Mindhárom fájl az eredeti export 13 oszlopát, UTF–16 kódolását, tabulátoros tagolását és tizedes vesszőjét használja; valódi ügyféladatot nem tartalmaz. Fájlonként 108 tétel, három nap, minden napon mindkét pénznemben 18 tétel van. A sorok fordított időrendben szerepelnek.

| Fájl a samples mappában | Dátumok | Fő ellenőrzés |
| --- | --- | --- |
| napokon_at.csv | 2026-09-18, 2026-09-19, 2026-09-21 | Többnapos import és hétvégi árfolyampótlás. |
| honapokon_at.csv | 2026-09-30, 2026-10-01, 2026-10-02 | Két hónap, megosztott 40. naptári hét. |
| eveken_at.csv | 2026-12-31, 2027-01-01, 2027-01-02 | Két éves riport, megosztott 2026-os 53. naptári hét. |

Minden nap/pénznem kombinációban van munkadíjas és munkadíj nélküli számla, új és bontott alkatrész, két munkadíjtétel ugyanazon számlán, SANY/DIAG/KJ, szállítás, Trailer, több csoportos és üres csoportú beszerzés, szállítólevél és leltár. A Szűrők csoport a kézi újalkatrész-csoport felvételét teszteli; vegyes kis- és nagybetűk, negatív korrekció és többdarabos tételek is szerepelnek. A leltárnak itt van tesztazonosítója, hogy az üres bizonylatszám ellenőrzése ne keveredjen a dátumhatárok vizsgálatába.

Az alábbi várt összegekben k a fájlon belüli nap időrendi sorszáma (1, 2, 3). A HUF és EUR nyers összeget külön adjuk meg, a végső beírás HUF + EUR × az adott naphoz rögzített árfolyam, majd egész forintos kerekítés. A Szűrők nincs felvéve további új csoportként.

| Cél | HUF nyers összeg | EUR nyers összeg |
| --- | ---: | ---: |
| Bontott alkatrész | 2800 × k | 28 × k |
| Új alkatrész | 7400 × k | 74 × k |
| Munkadíj | 4500 × k | 45 × k |
| Bontott műhely | 2000 × k | 20 × k |
| Új műhely | 6600 × k | 66 × k |
| Beszerzés | 2625 × k | 26,25 × k |
| Posta | 250 × k | 2,5 × k |
| Trailer | 1500 × k | 15 × k |
| SANY havi növekmény | 100 × k | 1 × k |
| DIAG havi növekmény | 200 × k | 2 × k |
| KJ havi növekmény | 300 × k | 3 × k |

A Szűrők felvétele az új alkatrész napi és műhelyösszegét egyaránt 700 × k HUF és 7 × k EUR összeggel növeli. A szállítólevél és leltár nem növeli a fenti célokat. Ugyanazon teljes fájl ismételt importjának minden növekménye nulla. A fájlok bizonylatai egymástól különböznek.

Árfolyampótlás ellenőrzése: egy hiányzó múltbeli naphoz ugyanazon hónap legközelebbi eredeti napi adata kerül; azonos távolságnál a korábbi. Ha a hónapban nincs forrásadat, kézi bevitel szükséges. Az eltárolt pótlás ismételt importkor új lekérés nélkül változatlan marad. A CSV-k árfolyamot nem tartalmaznak; az árfolyamforrást vagy a kézi bevitelt a teszt során biztosítjuk.

Rögzített ellenőrző árfolyamként például **400 HUF/EUR** használható (fiktív tesztérték, nem MNB-adat). Ezzel k=1 esetén a várt forintösszegek: bontott 14 000; új 37 000; munkadíj 22 500; bontott műhely 10 000; új műhely 33 000; beszerzés 13 125; posta 1 250; Trailer 7 500; SANY 500; DIAG 1 000; KJ 1 500. k=2 és k=3 esetén ezek kétszerese, illetve háromszorosa várható.