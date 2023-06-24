# Manual Sorter v1.2 jaoks


## Mis on OOP?

<p>
OOP ehk Objekt Orienteeritud Progarmeerimine tähendab et mingi koodi sees luuakse objekt millel on erinevad meetodid millega saab objekti see olevaid andmeid muuta või kätte saada.

näide reaalsest elust:
Oletame et meil on objekt mille nimi on auto.
kui vaadata peale siis me näeme lihtsalt autot aga kui me hakkame lähemalt vaatama siis auto koosneb ratastest ustest värvist jne jne. Programeerimises on täpselt sama lugu ma saame luua objekti millel anname nime ja määrame millest see objekt või asi koosneb.

</p>
#### Näide Sorteerimis koodist objektil Data
``` python
class Data:

    def __init__(self, code: str, postision: str, version: str = None) -> None:
        self.code = code
        self.version = version
        self.positsion = postision

    def __repr__(self) -> str:
        return f"Kood: {self.code} \n kogus: {self.amount} \n Versioon: {self.version} \n Positsioon: {self.positsion}"

    def getCode(self) -> str:
        return self.code
    
    
    def getVersion(self):
        return self.version
    
    def getPosition(self) -> str:
        return self.positsion
    
    def __eq__(self, other) -> bool:
        if isinstance(other, Data):
            return self.code == other.getCode  and self.version == other.getVersion and self.positsion == other.getPosition
        return False
```

<p>
 Siin me ütleme arvutile et on olemas objekt nimega Data milles me anname talle kaasa konstruktori kus ma määrame millest see objekt koosneb.
</p>

```python
    def __init__(self, code: str, postision: str, version: str = None) -> None:
        self.code = code
        self.version = version
        self.positsion = postision
```
See koodi jupp ütleb et selleks et luua objekti data tuleb ta sisse pana kood, positsioon ja versioon ning on oeldud et see objekt koosneb meetoditest mis koik tagastavad midagi nt meetod getCode() tagastab koodi mis objektile on antud.

#### Objekt DataBase

Nagu Nimi viitab siis see objekt on meie andme hoidja ehk siis see on objekt mis tegelikult on list mis koosneb Data objektidest. Näide: Kui meil on objekt auto ja autoparkla. ehk me saame vaadata mis on Autol aga meil on ka autoparkla kus siis hoitakse palju erinevad autosid ja me saame sealt pärast välja valida sellise nagu meil vaja on või saame autosid juurde panna.

##### Kood
```python
class DataBase:


    def __init__(self) -> None:
        self.info = []


    def  __repr__(self) -> str:
        return str(self.info)
    
    def add(self, item: Data):
        self.info.append(item)

    def getDataBase(self):
        return self.info
    
    def __add__(self, other):
        if isinstance(other, DataBase):
            self.info.extend(other.getDataBase())
        return self
    
    def getPositions(self) -> list:
        result = []
        for item in self.info:
            if item.getPosition() not in result:
                result.append(item.getPosition())
        return result


        # worker
    def getData(self, NameOfTheFIle: str) -> None:
        # Opens the file.
        with open(NameOfTheFIle, encoding='utf-8-sig',) as csv_file:

            csv_reader = list(csv.reader(csv_file, delimiter=";"))
            if "Versioon" in csv_reader[0]:
                versionIndex =  csv_reader[0].index("Versioon")
            else:
                versionIndex =  None
            koodIndex = csv_reader[0].index("Kood")
            kogusIndex = csv_reader[0].index("Kogus")
            positsioonIndex = csv_reader[0].index("Positsioon")
            for row in csv_reader[1:]:
                positsions = row[positsioonIndex].split(", ")
                for positsion in positsions:
                    if versionIndex == None:
                        item = Data(row[koodIndex], positsion)    
                    else:
                        item = Data(row[koodIndex], positsion, row[versionIndex])
                    self.add(item)
```
Siin klassis nagu näha ei taha konstruktor mitte midagi ehk siis me saame luua objekti database ilma et talle midagi annaks nagu naiteks me saame luue parkla ilma et seal peaks olema ühtegi autot. Küll aga on olemas meetodid add() ja getDatabase mis teevad seda mida nimed ütlevad. Esimesega saab panna database elemente ja teisega saab tagastada kogu sisu mis on databases

#### Meetod getData()

Nüüd meil on erinevalt Data klassist veel erinev meetod nimega getData millele tuleb anda üks teksi argument ja see on file nimi. See meetod on nagu põhimõtteliselt mul on antud hunnik juppe pane neist kokku autod ja hoia neid parklas.

```python
csv_reader = list(csv.reader(csv_file, delimiter=";"))
if "Versioon" in csv_reader[0]:
    versionIndex =  csv_reader[0].index("Versioon")
    else:
        versionIndex =  None
    koodIndex = csv_reader[0].index("Kood")
    kogusIndex = csv_reader[0].index("Kogus")
    positsioonIndex = csv_reader[0].index("Positsioon")
    for row in csv_reader[1:]:
        positsions = row[positsioonIndex].split(", ")
        for positsion in positsions:
            if versionIndex == None:
                item = Data(row[koodIndex], positsion)    
            else:
                item = Data(row[koodIndex], positsion, row[versionIndex])
            self.add(item)
```
siin me avame selle file ja ma sain teada et arvuti loeb igat ruudu vahe märki ";" tähena ehk siis delimiter=";" tähendab et iga kahe kasti vahel on see märk ja nii saab arvuti teha selle 2d listiks ehk list mis koosneb listidest kus list on kogu file ja siis iga lists olev list on siis rida sellest filest

Kuna me ei tea millises kastis on Kood, Kogus, Positsioon ega versioon siis ainult seda et esimesel real on nende tulpade nimed siis kuna me salvestasime selle csv file muutujana csv_reader siis kui kirjutada csv_reader[0] saame esimese rea filest. Järgmisena on meil muutujad koodIndex, versioonIndex, positsioonIndex. kasutades sisse ehitatud meetodit .index saame otsida listis olevat elementi ja tagastada selle koha index muutujasse. Sedasi saame teada millisele tulbale vastab mis väärtus.

#### 2 Tsüklit üksteise sees

kuna meil on vaja läbi käia kõik read ja meil on vaja ka teada mitu mingit asja on positsioonide järgi peame panema 2 tsüklit üksteise sisse

```python
for row in csv_reader[1:]:
    positsions = row[positsioonIndex].split(", ")
    for positsion in positsions:
        if versionIndex == None:
            item = Data(row[koodIndex], positsion)   
        else:
            item = Data(row[koodIndex], positsion, row[versionIndex])
        self.add(item)
```
Kood siis hakkab käima läbi igat rida jättes vahele esimese rea sest seal on kirjas tulba nimed. kuna me teame millisel kohal real on positsioonid siis me kasutame row[positsioonIndex].split(",") mis lõpuks annab list mille iga element on siis positsioon.
Nüüd meil on vaja luua objekte mis käivad selle toote kohta ehk siis teine tsükkel mis käib läbi seda listi kus on positsiooni nimed. Kuna versioon võib olla aga ei pea siis me vaatame kas versionIndex on olemas või ei ole. Kui on siis me loome objekt Data mis saab omale koodi positsiooni ja versioon kui ei ole siis versiooni ei saa. siis me lisame selle Database.

#### Kirjutame andmebaasid file ja toome välja erinevusi

Viimane osa on kuidas nüüd neid eelnvalt loodud objekte kasutada ja sealseid andmeid töödelda.

```Python
def magic(file_first: str, file_second: str, output: str) -> int:


    if not os.path.isfile(file_first):
        return -1
    if not os.path.isfile(file_second):
        return 0
    
    workbook = Workbook()
    sheet = workbook.active

    red = Font(color="FF0000")
    black = Font(color="000000")
    green = PatternFill(start_color="5bba75", end_color="5bba75", fill_type="solid")
    purple = PatternFill(start_color="c95da0", end_color="c95da0", fill_type="solid")
    
    file1 = DataBase()
    file1.getData(file_first)
    fileUnchanged1 = file1.getDataBase()

    file2 = DataBase()
    file2.getData(file_second)
    fileUnchanged2 = file2.getDataBase()

    positions = list(set((file1.getPositions() + file2.getPositions())))
    positions.sort()
    rows = [["Kood", "Versioon", "Kood", "Versioon", "Positsioon"]]

    for position in positions:
        foundFile1 = list(filter(lambda x: x.getPosition() == position, fileUnchanged1))
        foundFile2 = list(filter(lambda x: x.getPosition() == position, fileUnchanged2))
        if len(foundFile1) > 0 and len(foundFile2) > 0:
            row = [foundFile1[0].getCode(), foundFile1[0].getVersion()
                    , foundFile2[0].getCode(), foundFile2[0].getVersion(), foundFile2[0].getPosition()]
        if len(foundFile1) > 0 and len(foundFile2) == 0:
            row = [foundFile1[0].getCode(), foundFile1[0].getVersion(), "", "", foundFile1[0].getPosition()]
        if len(foundFile1) == 0 and len(foundFile2) > 0:
            row = ["", "", foundFile2[0].getCode(), foundFile2[0].getVersion(), foundFile2[0].getPosition()]
        rows.append(row)


    header_row = rows[0]
    for col_num, value in enumerate(header_row, start=1):
        cell = sheet.cell(row=1, column=col_num)
        if col_num == 1 or col_num == 2:
            cell.fill = green
        if col_num == 3 or col_num == 4:
            cell.fill = purple
        cell.value = value

    data = rows[1:]
    for row_num, row in enumerate(data, start=2):
        for col_num, value in enumerate(row, start=1):
            try:
                value = int(value)
            except ValueError:
                pass
            cell = sheet.cell(row=row_num, column=col_num)
            cell.value = value
            if col_num in [1, 3]:  # Columns 1 and 3
                if col_num == 1:
                    cell.fill = green
                    if(row[col_num - 1] != row[col_num + 1]):
                        cell.font = red
                if col_num == 3:
                    cell.fill = purple
                    if(row[col_num - 1] != row[col_num - 3]):
                        cell.font = red
            elif col_num in [2, 4]:  # Columns 2 and 4
                if col_num == 2:
                    cell.fill = green
                    if(row[col_num - 1] != row[col_num + 1]):
                        cell.font = red
                if col_num == 4:
                    cell.fill = purple
                    if(row[col_num - 1] != row[col_num - 3]):
                        cell.font = red
            elif col_num == 5:  # Column 5
                cell.font = black
    
    workbook.save(output)
    return 1
```

Ma ei oskand sellele funktsioonile muud nimeks panna sest see kuidas see töötab on ka maagia minu jaoks. Selle funktsiooni välja kutsumiseks tuleb talle anda esimese file nimi teise ja siis väljundi file nimi. Esimese asjana kontrollime kas mõlemad file on olemas või ei et kood ei satuks mingi errori otsa.

Nüüd me loome 2 Database 1 mis salvestab esimese file andmed ja teine mis salvestab teise. Siis me saame kätte kõikide objektide positsioonid ja paneme need ühte listi ja eemaldame duplikaadid. Kood koosneb mitmest väiksest sektsioonist mille ülesanded on järgnevad
1) Me saame kätte kõik võimalikud positsioonid
2) Käime läbi iga positsiooni ja vaatame kas me saame mõlemast andmebaasist selle koha peale vaste ja lisame oma 2d list uue listi mis koosneb mingi rea andmetest.
3) hakkame iga rida kirjutama exceli file ja kuna me teame et nüüd mis järjekorras on asjad sest ennem said need ise tehtud siis teame täpselt mis indexil midagi asub vaatame kas koodi ja versiooni indexitel olevad väärtused on samad või ei kui ei ole siis me paneme selle exceli file punasena kui mitte siis mustana.

Edu selle koodi mõistmisega kallis haha