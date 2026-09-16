// Query - ImportJSON вЂ” CORE, generic ETL-РїР°Р№РїР»Р°Р№РЅ (Architecture Core В§5, В§5.1).
// Р’С‹Р·С‹РІР°РµС‚ РїРѕ С„РёРєСЃРёСЂРѕРІР°РЅРЅС‹Рј РёРјРµРЅР°Рј РєРѕРЅС‚СЂР°РєС‚Р°: fnNormalizeFields, fnComputeKey,
// fnDedupByKey, fnComputeGroupMetrics (СЂРµР°Р»РёР·Р°С†РёСЏ вЂ” Content Spec РЅР°РїСЂР°РІР»РµРЅРёСЏ)
// Рё fnUpsert (Core).
//
// prmSourcePath вЂ” РёРјРµРЅРѕРІР°РЅРЅС‹Р№ РїР°СЂР°РјРµС‚СЂ (С‚РµРєСЃС‚РѕРІС‹Р№), Р·Р°РґР°С‘С‚СЃСЏ РёР· VBA РїРµСЂРµРґ Refresh (modMain.bas).
// "tbDATA" вЂ” РёРјСЏ С†РµР»РµРІРѕР№ С‚Р°Р±Р»РёС†С‹ Р»РёСЃС‚Р°, С„РёРєСЃРёСЂРѕРІР°РЅРѕ РєРѕРЅС‚СЂР°РєС‚РѕРј Core (Architecture Core В§5.1).
//
// v7.4 (11.09.2026):
//   R-2 - prmSourcePath РїСЂРёРЅРёРјР°РµС‚ РџРђРџРљРЈ: РїРѕР»РЅР°СЏ РїРµСЂРµСЃР±РѕСЂРєР° Р·Р° РѕРґРёРЅ РїСЂРѕС…РѕРґ,
//         Р±РµР· РїРµСЂРµС‡РёС‚С‹РІР°РЅРёСЏ tbDATA (Р·Р°РјРµСЂ: РѕРґРёРЅ С„Р°Р№Р» С‡РµСЂРµР· upsert - 5 С‡ 12 РјРёРЅ).
//
// v8.2 (13.09.2026) - РўР— v1.0: РѕРґРЅРѕРєСЂР°С‚РЅРѕРµ С‡С‚РµРЅРёРµ РїР°РїРєРё (Р±СѓС„РµСЂРёР·Р°С†РёСЏ Source)
//   Рё РѕС‡РµСЂРµРґСЊ С„Р°Р№Р»РѕРІ РїРѕ РґР°С‚Рµ/РІСЂРµРјРµРЅРё РёР· РёРјРµРЅРё (fallback - РїРѕ Name).
//
// v7.3 (09.09.2026), РѕРґРЅР° РїСЂР°РІРєР°:
//   R-1 вЂ” РґРѕР±Р°РІР»РµРЅ С„РёРЅР°Р»СЊРЅС‹Р№ С€Р°Рі СЂРµС‚РµРЅС€РЅР° (РёРЅСЃС‚СЂСѓРєС†РёСЏ РїСЂРѕРµРєС‚Р° В§4): СЃС‚СЂРѕРєРё СЃС‚Р°СЂС€Рµ
//         DATA/KEEP_WEEKS РЅРµРґРµР»СЊ СѓРґР°Р»СЏСЋС‚СЃСЏ РёР· СЂРµР·СѓР»СЊС‚Р°С‚Р°. Р“Р»СѓР±РёРЅР° С‡РёС‚Р°РµС‚СЃСЏ РЅРѕРІС‹Рј
//         Р·Р°РїСЂРѕСЃРѕРј qKeepWeeks (Р»РёСЃС‚ Variable). РџСЂР°РІРєР° Core, РєСЂРёС‚РµСЂРёР№ СЃРѕР±Р»СЋРґС‘РЅ:
//         РѕРіСЂР°РЅРёС‡РµРЅРёРµ СЂРѕСЃС‚Р° С‚Р°Р±Р»РёС†С‹ РЅСѓР¶РЅРѕ Р›Р®Р‘РћРњРЈ РѕС‚С‡С‘С‚Сѓ РЅР° СЌС‚РѕР№ РїР»Р°С‚С„РѕСЂРјРµ, Р° РЅРµ РњРўРћ.
//         РџСЂР°РІРёР»Рѕ РІС‹Р±РѕСЂР° РѕРїРѕСЂРЅРѕР№ РґР°С‚С‹ (status_date, РёРЅР°С‡Рµ date) СЃРїРµС†РёС„РёС‡РЅРѕ РњРўРћ Рё Р¶РёРІС‘С‚
//         РІ С‚РµР»Рµ С€Р°РіР°; РїСЂРё РїРµСЂРµРЅРѕСЃРµ РїР»Р°С‚С„РѕСЂРјС‹ РїСЂР°РІРёС‚СЃСЏ С‚РѕР»СЊРєРѕ РѕРЅРѕ.
//
// v7 (06.09.2026), РѕРґРЅР° РїСЂР°РІРєР°:
//   C-1 вЂ” РІ РєРѕРЅРІРµР№РµСЂ РґРѕР±Р°РІР»РµРЅ С€Р°Рі fnDedupByKey РјРµР¶РґСѓ fnComputeKey Рё fnComputeGroupMetrics.
//         РџСЂР°РІРєР° Core, РєСЂРёС‚РµСЂРёР№ СЃРѕР±Р»СЋРґС‘РЅ: fnUpsert СЃРєР»РµРёРІР°РµС‚ СЂРµР·СѓР»СЊС‚Р°С‚ РєР°Рє
//         Table.Combine({NewBuf, Cleaned}) Рё РќР• РїСЂРѕРІРµСЂСЏРµС‚ СѓРЅРёРєР°Р»СЊРЅРѕСЃС‚СЊ Key РІ newData вЂ”
//         Р·РЅР°С‡РёС‚ РіР°СЂР°РЅС‚РёСЂРѕРІР°С‚СЊ РµС‘ РѕР±СЏР·Р°РЅ РєРѕРЅРІРµР№РµСЂ, Рё РѕР±СЏР·Р°РЅ РґР»СЏ Р›Р®Р‘РћР“Рћ РѕС‚С‡С‘С‚Р° РЅР° СЌС‚РѕР№
//         РїР»Р°С‚С„РѕСЂРјРµ, Р° РЅРµ С‚РѕР»СЊРєРѕ РґР»СЏ РњРўРћ. Р’ Core Р»РµР¶РёС‚ С‚РѕР»СЊРєРѕ РІС‹Р·РѕРІ РїРѕ С„РёРєСЃРёСЂРѕРІР°РЅРЅРѕРјСѓ
//         РёРјРµРЅРё; СЃР°РјРѕ РїСЂР°РІРёР»Рѕ РїРѕР±РµРґРёС‚РµР»СЏ (РґР»СЏ РњРўРћ вЂ” РјР°РєСЃРёРјР°Р»СЊРЅС‹Р№ status_date) Р¶РёРІС‘С‚
//         РІ СЂРµР°Р»РёР·Р°С†РёРё РЅР°РїСЂР°РІР»РµРЅРёСЏ, СЃРј. fnDedupByKey.pq.
//         РњРµСЃС‚Рѕ С€Р°РіР° СЃСѓС‰РµСЃС‚РІРµРЅРЅРѕ: РґРµРґСѓРїР»РёРєР°С†РёСЏ Р”Рћ fnComputeGroupMetrics РѕСЃС‚Р°РІР»СЏРµС‚ РЅР° РїР°СЂСѓ
//         (number, direction) РЅРµ Р±РѕР»СЊС€Рµ РѕРґРЅРѕР№ Р·Р°РїРёСЃРё РєР°Р¶РґРѕРіРѕ СЃС‚Р°С‚СѓСЃР°, Рё РґРѕРїСѓС‰РµРЅРёРµ min/max
//         РІ СЂР°СЃС‡С‘С‚Рµ deltaHours РІС‹СЂРѕР¶РґР°РµС‚СЃСЏ.
//
// v5 (02.09.2026), РґРІРµ РїСЂР°РІРєРё, РѕР±Рµ СЃРІСЏР·Р°РЅС‹ СЃ В«С„Р°Р№Р» РЅРµ Р·Р°РіСЂСѓР¶Р°РµС‚СЃСЏВ»:
//   L-1 вЂ” Json.Document РїРѕР»СѓС‡Р°РµС‚ РєРѕРґРѕРІСѓСЋ СЃС‚СЂР°РЅРёС†Сѓ 65001 (UTF-8). Р‘РµР· РІС‚РѕСЂРѕРіРѕ Р°СЂРіСѓРјРµРЅС‚Р°
//         РґРІРёР¶РѕРє СѓРіР°РґС‹РІР°РµС‚ РєРѕРґРёСЂРѕРІРєСѓ: РІС‹РіСЂСѓР·РєР° РІ Windows-1251 Р»РѕРјР°РµС‚ РїР°СЂСЃРµСЂ РЅР° РїРµСЂРІРѕР№ Р¶Рµ
//         РєРёСЂРёР»Р»РёС†Рµ, UTF-8 СЃ BOM вЂ” РЅР° РїРµСЂРІРѕРј С‚РѕРєРµРЅРµ. Р’ РіСЂР°Р±Р»СЏС… РїСЂРѕРµРєС‚Р° СЌС‚Рѕ СѓР¶Рµ Р·Р°РїРёСЃР°РЅРѕ,
//         РЅРѕ РІ РєРѕРґРµ СѓРєР°Р·Р°РЅРёСЏ РєРѕРґРёСЂРѕРІРєРё РЅРµ Р±С‹Р»Рѕ.
//   L-2 вЂ” С‡С‚РµРЅРёРµ СЃСѓС‰РµСЃС‚РІСѓСЋС‰РµРіРѕ Р»РёСЃС‚Р° РІС‹РЅРµСЃРµРЅРѕ РІ РѕС‚РґРµР»СЊРЅС‹Р№ Р·Р°РїСЂРѕСЃ qExistingData.
//         Power Query Р·Р°РїСЂРµС‰Р°РµС‚ РѕРґРЅРѕРјСѓ Р·Р°РїСЂРѕСЃСѓ РѕРґРЅРѕРІСЂРµРјРµРЅРЅРѕ РѕР±СЂР°С‰Р°С‚СЊСЃСЏ Рє С„Р°Р№Р»РѕРІРѕР№ СЃРёСЃС‚РµРјРµ
//         (File.Contents) Рё Рє РєРЅРёРіРµ (Excel.CurrentWorkbook) вЂ” СЌС‚Рѕ Formula.Firewall:
//         "references other queries or steps, so it may not directly access a data source".
//         (!) РћРґРЅР° СЌС‚Р° РїСЂР°РІРєР° firewall РЅРµ СЃРЅРёРјР°РµС‚: РЅСѓР¶РЅС‹ РµС‰С‘ РґРІРµ РІРµС‰Рё, СЃРј. AUDIT_src_v2.md,
//            СЂР°Р·РґРµР» В«Р‘Р»РѕРєРµСЂ 1В» (prmSourcePath РґРѕР»Р¶РµРЅ РѕСЃС‚Р°С‚СЊСЃСЏ РџРђР РђРњР•РўР РћРњ + СЂР°Р·РѕРІРѕ РІС‹РєР»СЋС‡РёС‚СЊ
//            СѓСЂРѕРІРЅРё РєРѕРЅС„РёРґРµРЅС†РёР°Р»СЊРЅРѕСЃС‚Рё РІ РЅР°СЃС‚СЂРѕР№РєР°С… Excel).
//
// v3.1 (СЂРµРІСЊСЋ 24.08.2026) вЂ” Р±РµР· РёР·РјРµРЅРµРЅРёР№ РІ v5:
//   A-3/P2-4 вЂ” fnNormalizeFields РІС‹Р·С‹РІР°РµС‚СЃСЏ РєР°Рє table -> table.
//   P1-7     вЂ” Table.Buffer РїРµСЂРµРґ РІС‹РіСЂСѓР·РєРѕР№ (Р·Р°РїСЂРѕСЃ СЃР°РјРѕСЃСЃС‹Р»Р°СЋС‰РёР№СЃСЏ).
//   P1-7b    вЂ” РїРµСЂРІС‹Р№ Р·Р°РїСѓСЃРє РЅРµ РѕРїРёСЂР°РµС‚СЃСЏ РЅР° fnUpsert(X, X).
let
    // 65001 вЂ” РєРѕРґРѕРІР°СЏ СЃС‚СЂР°РЅРёС†Р° РёСЃС‚РѕС‡РЅРёРєР°, Р° РЅРµ В«РєРѕРґРёСЂРѕРІРєР° РІС‹РІРѕРґР°В». UTF-8 СЃ BOM С‚РѕР¶Рµ РїРѕРЅРёРјР°РµС‚СЃСЏ.
    // R-2 (11.09.2026): prmSourcePath РјРѕР¶РµС‚ СѓРєР°Р·С‹РІР°С‚СЊ РЅР° Р¤РђР™Р› РёР»Рё РЅР° РџРђРџРљРЈ.
    // РџР°РїРєР° - СЂРµР¶РёРј РїРѕР»РЅРѕР№ РїРµСЂРµСЃР±РѕСЂРєРё: С‡РёС‚Р°СЋС‚СЃСЏ РІСЃРµ *.json СЃСЂР°Р·Сѓ, РѕРґРЅРёРј РїСЂРѕС…РѕРґРѕРј.
    // Р—Р°С‡РµРј: upsert РїРµСЂРµС‡РёС‚С‹РІР°РµС‚ РІРµСЃСЊ tbDATA СЃ Р»РёСЃС‚Р° РїСЂРё РєР°Р¶РґРѕРј С„Р°Р№Р»Рµ, Рё РЅР°
    // 255 С‚С‹СЃ. СЃС‚СЂРѕРє РѕРґРёРЅ С„Р°Р№Р» Р·Р°РЅРёРјР°Р» 5 С‡ 12 РјРёРЅ (Р·Р°РјРµСЂ 11.09.2026). РџСЏС‚СЊ С„Р°Р№Р»РѕРІ
    // РїРѕ С‚Р°РєРѕР№ СЃС…РµРјРµ - СЃСѓС‚РєРё. РћРґРЅРёРј РїСЂРѕС…РѕРґРѕРј РїРµСЂРµС‡РёС‚С‹РІР°РЅРёРµ РЅРµ РЅСѓР¶РЅРѕ РІРѕРІСЃРµ.
    // Р РµР¶РёРј РѕРґРЅРѕРіРѕ С„Р°Р№Р»Р° (РєРЅРѕРїРєР° В«Р—Р°РіСЂСѓР·РёС‚СЊВ») РЅРµ РјРµРЅСЏР»СЃСЏ.
    IsFolder = not Text.EndsWith(Text.Lower(Text.Trim(prmSourcePath)), ".json"),

    // РР·РІР»РµС‡РµРЅРёРµ РґР°С‚С‹/РІСЂРµРјРµРЅРё РёР· РёРјРµРЅРё С„Р°Р№Р»Р° (С€Р°Р±Р»РѕРЅ \d{8}_\d{6}): РїРѕР·РёС†РёСЏ "_",
    // 8 С†РёС„СЂ РґРѕ Рё 6 РїРѕСЃР»Рµ. РќРµ РЅР°Р№РґРµРЅРѕ - "".
    ExtractStamp = (name as text) as text =>
        let
            positions = Text.PositionOfAny(name, {"_"}, Occurrence.All),
            valid = List.Select(positions, each
                _ >= 8 and
                Text.Length(Text.Select(Text.Range(name, _ - 8, 8), {"0".."9"})) = 8 and
                Text.Length(Text.Select(Text.Range(name, _ + 1, 6), {"0".."9"})) = 6
            )
        in if List.IsEmpty(valid) then "" else Text.Range(name, List.Min(valid) - 8, 15),

    // РљР»СЋС‡ СЃРѕСЂС‚РёСЂРѕРІРєРё РѕС‡РµСЂРµРґРё (РўР— v1.0, Рї.3.2): РґР°С‚Р°/РІСЂРµРјСЏ РёР· РёРјРµРЅРё;
    // РЅРµ СѓРґР°Р»РѕСЃСЊ РёР·РІР»РµС‡СЊ - fallback РЅР° РёРјСЏ С„Р°Р№Р»Р°.
    SortKey = (name as text) as text => let s = ExtractStamp(name) in if s = "" then name else s,

    FolderFiles =
        if not IsFolder then #table({}, {})
        else Table.SelectRows(
                Folder.Files(prmSourcePath),
                each Text.EndsWith(Text.Lower([Name]), ".json")
             ),

    // Р‘СѓС„РµСЂРёР·Р°С†РёСЏ (РўР— v1.0, Рї.3.1): List.Buffer РґР»СЏ List.Combine - РґР°РЅРЅС‹Рµ РёР· data
    // С‡РёС‚Р°СЋС‚СЃСЏ РµРґРёРЅРѕР¶РґС‹, Р° РЅРµ РїРµСЂРµСЃС‡РёС‚С‹РІР°СЋС‚СЃСЏ РЅР° СЃР»РµРґСѓСЋС‰РёС… С€Р°РіР°С… (SampleFields/Expanded).
    // (!) List.Buffer, Р° РЅРµ Table.Buffer: Source - РЎРџРРЎРћРљ Р·Р°РїРёСЃРµР№ (List.Combine), Р°
    //     Table.Buffer С‚СЂРµР±СѓРµС‚ С‚Р°Р±Р»РёС†Сѓ Рё СѓСЂРѕРЅРёР» Р±С‹ СЂРµР¶РёРј РїР°РїРєРё РЅР° С‚РёРїРµ Р°СЂРіСѓРјРµРЅС‚Р°.
    // РЎРѕСЂС‚РёСЂРѕРІРєР° (РўР— v1.0, Рї.3.2): РѕС‡РµСЂРµРґСЊ РїРѕ РґР°С‚Рµ/РІСЂРµРјРµРЅРё РёР· РёРјРµРЅРё С„Р°Р№Р»Р°; fallback - Name.
    Source =
        if not IsFolder then Json.Document(File.Contents(prmSourcePath), 65001)
        else List.Buffer(
                List.Combine(
                    List.Transform(
                        Table.ToRecords(Table.Sort(FolderFiles, (a, b) => Value.Compare(SortKey(a[Name]), SortKey(b[Name])))),
                        each Json.Document([Content], 65001)
                    )
                )
             ),

    // РќР°Р±РѕСЂ РєРѕР»РѕРЅРѕРє РІС‹РІРѕРґРёС‚СЃСЏ РїРѕ РѕР±СЉРµРґРёРЅРµРЅРёСЋ РїРѕР»РµР№ РІС‹Р±РѕСЂРєРё Р·Р°РїРёСЃРµР№, Р° РЅРµ РїРѕ РїРµСЂРІРѕР№ Р·Р°РїРёСЃРё:
    // 1РЎ РјРѕР¶РµС‚ РѕРїСѓСЃС‚РёС‚СЊ null-РїРѕР»Рµ (defekt_type/defect_desc) РІ РѕС‚РґРµР»СЊРЅС‹С… Р·Р°РїРёСЃСЏС….
    // (!) Р”РѕРїСѓС‰РµРЅРёРµ (СЃРѕС…СЂР°РЅРµРЅРѕ РёР· v3.1): РІС‹Р±РѕСЂРєР° вЂ” РїРµСЂРІС‹Рµ 1000 Р·Р°РїРёСЃРµР№. Р•СЃР»Рё РїРѕР»Рµ РІРїРµСЂРІС‹Рµ
    //    РІСЃС‚СЂРµС‡Р°РµС‚СЃСЏ РїРѕР·Р¶Рµ, РѕРЅРѕ РјРѕР»С‡Р° РЅРµ РїРѕРїР°РґС‘С‚ РІ С‚Р°Р±Р»РёС†Сѓ. Р”РёР°РіРЅРѕСЃС‚РёРєР° вЂ” СЃРј. qDiagImport.
    SampleFields = List.Distinct(
        List.Combine(List.Transform(List.FirstN(Source, 1000), each Record.FieldNames(_)))
    ),
    Expanded = Table.FromRecords(Source, SampleFields, MissingField.UseNull),

    // РќРѕСЂРјР°Р»РёР·Р°С†РёСЏ + С‚РёРїРёР·Р°С†РёСЏ РїРѕР»РµР№ РЅР°РїСЂР°РІР»РµРЅРёСЏ (РєРѕРЅС‚СЂР°РєС‚ table -> table).
    Normalized = fnNormalizeFields(Expanded),

    WithKey = Table.AddColumn(Normalized, "Key", each fnComputeKey(_), type text),

    // C-1: СѓРЅРёРєР°Р»СЊРЅРѕСЃС‚СЊ Key РІ РїСЂРµРґРµР»Р°С… РѕРґРЅРѕРіРѕ С„Р°Р№Р»Р°. РћР±СЏР·Р°С‚РµР»СЊРЅР° РїРµСЂРµРґ fnUpsert
    // Рё РїРµСЂРµРґ СЂР°СЃС‡С‘С‚РѕРј РґРµР»СЊС‚ вЂ” СЃРј. С€Р°РїРєСѓ Рё fnDedupByKey.pq.
    Deduped = fnDedupByKey(WithKey),

    WithGroups = fnComputeGroupMetrics(Deduped),

    NewData = WithGroups,

    // РЎСѓС‰РµСЃС‚РІСѓСЋС‰РёРµ РґР°РЅРЅС‹Рµ Р»РёСЃС‚Р° вЂ” РѕС‚РґРµР»СЊРЅС‹Р№ Р·Р°РїСЂРѕСЃ (СЃРј. qExistingData.pq),
    // РіСЂСѓР·РёС‚СЃСЏ РєР°Рє В«РўРѕР»СЊРєРѕ СЃРѕР·РґР°С‚СЊ РїРѕРґРєР»СЋС‡РµРЅРёРµВ».
    ExistingData = #table({}, {}),

    Upserted =
        if IsFolder then NewData
        else if Table.RowCount(ExistingData) = 0 then NewData
        else if not List.Contains(Table.ColumnNames(ExistingData), "Key") then NewData
        else fnUpsert(NewData, ExistingData),

    // Р‘СѓС„РµСЂРёР·Р°С†РёСЏ РѕР±СЏР·Р°С‚РµР»СЊРЅР°: РЅРёР¶Рµ Upserted С‡РёС‚Р°РµС‚СЃСЏ С‚СЂРёР¶РґС‹ (РґРІР° СЃРєР°РЅР° РєРѕР»РѕРЅРѕРє + С„РёР»СЊС‚СЂ).
    // Р‘РµР· Table.Buffer РґРІРёР¶РѕРє РїРµСЂРµСЃС‡РёС‚Р°РµС‚ РІРµСЃСЊ upsert РЅР° РєР°Р¶РґРѕРµ РѕР±СЂР°С‰РµРЅРёРµ.
    UpsertedBuf = Table.Buffer(Upserted),

    // ============================ R-1: Р Р•РўР•РќРЁРќ tbDATA ============================
    // РРЅСЃС‚СЂСѓРєС†РёСЏ РїСЂРѕРµРєС‚Р° В§4. Р“Р»СѓР±РёРЅР° - С‚РѕР»СЊРєРѕ СЃ Р»РёСЃС‚Р° Variable, СЃРј. qKeepWeeks.pq.
    // KeepWeeks = 0 -> С€Р°Рі РЅРµ РІС‹РїРѕР»РЅСЏРµС‚СЃСЏ, С‚Р°Р±Р»РёС†Р° РІРѕР·РІСЂР°С‰Р°РµС‚СЃСЏ РєР°Рє РµСЃС‚СЊ.
    KeepWeeks = 0,

    // РџСЂРёРІРµРґРµРЅРёРµ Рє datetime С‚РµСЂРїРёРјРѕРµ, Рё СЌС‚Рѕ РЅРµ РїРµСЂРµСЃС‚СЂР°С…РѕРІРєР°: РЅРѕРІС‹Рµ СЃС‚СЂРѕРєРё С‚РёРїРёР·РёСЂРѕРІР°РЅС‹
    // РІ fnNormalizeFields, Р° СЃС‚Р°СЂС‹Рµ РїСЂРёС…РѕРґСЏС‚ РѕР±СЂР°С‚РЅРѕ С‡РµСЂРµР· qExistingData РёР· СЏС‡РµРµРє Р»РёСЃС‚Р°
    // РєР°Рє any/text (СЃРј. Р±Р»РѕРє В«РџСЂРѕ С‚РёРїС‹В» РІ qExistingData.pq). РЎС‚СЂРѕРіРѕРµ DateTime.From
    // СѓСЂРѕРЅРёР»Рѕ Р±С‹ Р·Р°РїСЂРѕСЃ РЅР° РїРµСЂРІРѕР№ Р¶Рµ СЃС‚Р°СЂРѕР№ СЃС‚СЂРѕРєРµ.
    ToDT = (v) as nullable datetime =>
        if IsFolder then
            if v = null then null
            else if Value.Is(v, type datetime) then v
            else if Value.Is(v, type date) then DateTime.From(v)
            else null
        else
            if v = null then null
            else if Value.Is(v, type datetime) then v
            else if Value.Is(v, type date) then DateTime.From(v)
            else if Value.Is(v, type number) then (try DateTime.From(v) otherwise null)
            else
                let s = Text.Trim(Text.From(v))
                in  if s = "" then null
                    else try DateTime.From(s)
                         otherwise try DateTime.From(s, "ru-RU")
                         otherwise try DateTime.From(s, "en-US")
                         otherwise null,

    // РћРїРѕСЂРЅР°СЏ РґР°С‚Р° СЃС‚СЂРѕРєРё: РґР°С‚Р° РїРѕРґРїРёСЃР°РЅРёСЏ, Р° РµСЃР»Рё РµС‘ РЅРµС‚ (arm = В«РќР• РџРћР”РџРРЎРђРќРћВ»,
    // status_date РїСѓСЃС‚ Сѓ 38% РІС‹РіСЂСѓР·РєРё) - РґР°С‚Р° СЃРѕР·РґР°РЅРёСЏ РЅР°СЂСЏРґР°. date РЅРµРїСѓСЃС‚ РІСЃРµРіРґР°.
    RowAnchor = (r as record) as nullable datetime =>
        if IsFolder then
            let s = Record.FieldOrDefault(r, "status_date", null),
                d = Record.FieldOrDefault(r, "date", null)
            in  if s <> null and d <> null then (if s > d then s else d)
                else if s <> null then s
                else d
        else
            let s = ToDT(Record.FieldOrDefault(r, "status_date", null)),
                d = ToDT(Record.FieldOrDefault(r, "date", null))
            in  if s <> null and d <> null then List.Max({s, d})
                else if s <> null then s
                else d,

    ColDates = (t as table, col as text) as list =>
        if not List.Contains(Table.ColumnNames(t), col) then {}
        else if IsFolder then List.RemoveNulls(Table.Column(t, col))
        else List.RemoveNulls(List.Transform(Table.Column(t, col), each ToDT(_))),

    Anchors = List.Combine({ColDates(UpsertedBuf, "status_date"), ColDates(UpsertedBuf, "date")}),
    MaxDt   = if List.IsEmpty(Anchors) then null else List.Max(Anchors),

    // РћС‚СЃРµС‡РєР° - РїРѕРЅРµРґРµР»СЊРЅРёРє РЅРµРґРµР»Рё, РѕС‚СЃС‚РѕСЏС‰РµР№ РѕС‚ СЃР°РјРѕР№ СЃРІРµР¶РµР№ РЅРµРґРµР»Рё С‚Р°Р±Р»РёС†С‹
    // РЅР° KeepWeeks-1: РїСЂРё KeepWeeks = 1 РѕСЃС‚Р°С‘С‚СЃСЏ РѕРґРЅР° С‚РµРєСѓС‰Р°СЏ РЅРµРґРµР»СЏ.
    Cutoff =
        if KeepWeeks < 1 or MaxDt = null then null
        else Date.AddDays(Date.StartOfWeek(DateTime.Date(MaxDt), Day.Monday), -7 * (KeepWeeks - 1)),

    // РЎС‚СЂРѕРєР° Р±РµР· СЂР°СЃРїРѕР·РЅР°РЅРЅРѕР№ РѕРїРѕСЂРЅРѕР№ РґР°С‚С‹ РЎРћРҐР РђРќРЇР•РўРЎРЇ. РўР°РєР°СЏ СЃС‚СЂРѕРєР° РЅРµ РґРѕР»Р¶РЅР° РїРѕСЏРІР»СЏС‚СЊСЃСЏ
    // (date РЅРµРїСѓСЃС‚ РІСЃРµРіРґР°), РЅРѕ РјРѕР»С‡Р° СѓРґР°Р»СЏС‚СЊ С‚Рѕ, С‡С‚Рѕ РЅРµ СЃРјРѕРіР»Рё СЂР°Р·РѕР±СЂР°С‚СЊ, РЅРµР»СЊР·СЏ.
    Retained =
        if Cutoff = null then UpsertedBuf
        else Table.SelectRows(UpsertedBuf, each
                let a = RowAnchor(_) in a = null or DateTime.Date(a) >= Cutoff)
in
    #table({"Rows"}, {{ if Value.Is(Source, type table) then Table.RowCount(Source) else if Value.Is(Source, type list) then List.Count(Source) else 1 }})
