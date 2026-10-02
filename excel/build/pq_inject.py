"""Nhúng các query Power Query (.m) vào file xlsx dựng bằng openpyxl.

Tạo: customXml DataMashup (định dạng MS-QDEFF), xl/connections.xml, xl/queryTables/*, gắn bảng KQ_* thành bảng
kết quả của query (tableType="queryTable"), đánh dấu công thức mảng động (dynamic array) cho Excel 365.
Định dạng này không kiểm thử được bằng Excel trong môi trường dựng; nếu Excel báo lỗi phần query, dùng các file
power_query/*.m để tạo lại query bằng tay (xem README).
"""
import base64
import io
import re
import struct
import uuid
import zipfile
from xml.sax.saxutils import escape

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

META_DA = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
           '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
           'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
           'coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
           '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
           '</ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')


def _zip(files):
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            z.writestr(name, data)
    return b.getvalue()


def _guid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def data_mashup(queries, loaded, counts):
    section = "section Section1;\r\n"
    for name, code in queries.items():
        section += f"\r\nshared {name} = {code.strip()};\r\n"
    ct = ('<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="xml" ContentType="text/xml" /><Default Extension="m" ContentType="application/x-ms-m" /></Types>')
    pkg_xml = ('<?xml version="1.0" encoding="utf-8"?><Package xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
               'xmlns:xsd="http://www.w3.org/2001/XMLSchema"><Version>2.0.0.0</Version><MinVersion>2.21.0.0</MinVersion>'
               '<Culture>vi-VN</Culture><SafeCombine>true</SafeCombine><Items /></Package>')
    bom = b"\xef\xbb\xbf"
    parts = _zip([("[Content_Types].xml", bom + ct.encode()), ("Config/Package.xml", bom + pkg_xml.encode()),
                  ("Formulas/Section1.m", section.encode("utf-8"))])
    perm = (bom + b'<?xml version="1.0" encoding="utf-8"?><PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
            b'xmlns:xsd="http://www.w3.org/2001/XMLSchema"><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>'
            b'<FirewallEnabled>true</FirewallEnabled><WorkbookGroupType xsi:nil="true" /></PermissionList>')
    items = ['<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries>'
             '<Entry Type="Relationships" Value="sAAAAAA==" /></StableEntries></Item>']
    for name, code in queries.items():
        ham = code.lstrip().startswith("(") or "=>" in code.split("let")[0]
        kq = "sFunction" if ham else ("sRecord" if name == "ThamSo" else "sTable")
        e = [("IsPrivate", "l0"), ("ResultType", kq), ("BufferNextRefresh", "l1"), ("NavigationStepName", "sNavigation"),
             ("QueryID", "s" + str(uuid.uuid4()))]
        if name in loaded:
            cols = "[" + ",".join('"' + c + '"' for c in loaded[name]) + "]"
            e += [("FillEnabled", "l1"), ("FillObjectType", "sTable"), ("FillToDataModelEnabled", "l0"),
                  ("FillTarget", "s" + name), ("FilledCompleteResultToWorksheet", "l1"), ("AddedToDataModel", "l0"),
                  ("FillCount", f"l{counts[name]}"), ("FillErrorCode", "sUnknown"), ("FillErrorCount", "l0"),
                  ("FillLastUpdated", "d2026-10-02T00:00:00.0000000Z"), ("FillColumnNames", "s" + cols),
                  ("FillStatus", "sComplete")]
        else:
            e += [("FillEnabled", "l0"), ("FillObjectType", "sConnectionOnly"), ("FillToDataModelEnabled", "l0")]
        ent = "".join(f'<Entry Type="{t}" Value="{escape(v, {chr(34): "&quot;"})}" />' for t, v in e)
        items.append(f"<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/{name}</ItemPath></ItemLocation>"
                     f"<StableEntries>{ent}</StableEntries></Item>")
    meta_xml = (bom + ('<?xml version="1.0" encoding="utf-8"?><LocalPackageMetadataFile xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
                       'xmlns:xsd="http://www.w3.org/2001/XMLSchema"><Items>' + "".join(items) + "</Items></LocalPackageMetadataFile>").encode())
    meta_content = _zip([("[Content_Types].xml", bom + ct.encode())])
    meta = struct.pack("<I", 0) + struct.pack("<I", len(meta_xml)) + meta_xml + struct.pack("<I", len(meta_content)) + meta_content
    blob = (struct.pack("<I", 0) + struct.pack("<I", len(parts)) + parts + struct.pack("<I", len(perm)) + perm
            + struct.pack("<I", len(meta)) + meta + struct.pack("<I", 0))
    xml = ('<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
           + base64.b64encode(blob).decode() + "</DataMashup>")
    return b"\xff\xfe" + xml.encode("utf-16-le"), section


def inject(src, dst, queries, loaded):
    zin = zipfile.ZipFile(src)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist()}
    order = [i.filename for i in zin.infolist()]

    # --- bảng: tìm file theo displayName
    table_file = {}
    for f, data in files.items():
        if f.startswith("xl/tables/table") and f.endswith(".xml"):
            m = re.search(rb'displayName="([^"]+)"', data)
            table_file[m.group(1).decode()] = f
    # --- sheet: tên, chỉ số, file
    wbx = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx)
    rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})

    counts, defined, ct_add, rel_add = {}, [], [], []
    conns = []
    for n, (q, cols) in enumerate(loaded.items(), 1):
        tf = table_file[q]
        t = files[tf].decode()
        ref = re.search(r'ref="([^"]+)"', t).group(1)
        r1, r2 = [int(re.sub(r"[A-Z$]", "", x)) for x in ref.split(":")]
        counts[q] = r2 - r1
        t = t.replace("<table ", '<table tableType="queryTable" ', 1)
        for j, c in enumerate(cols, 1):
            t = re.sub(rf'<tableColumn id="{j}" name="{re.escape(c)}"', f'<tableColumn id="{j}" uniqueName="{j}" name="{c}" queryTableFieldId="{j}"', t)
        files[tf] = t.encode()
        qt = f"xl/queryTables/queryTable{n}.xml"
        fields = "".join(f'<queryTableField id="{j}" name="{c}" tableColumnId="{j}"/>' for j, c in enumerate(cols, 1))
        files[qt] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<queryTable xmlns="{MAIN}" '
                     f'name="ExternalData_{n}" connectionId="{n}" autoFormatId="16" applyNumberFormats="0" applyBorderFormats="0" '
                     f'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="0">'
                     f'<queryTableRefresh nextId="{len(cols) + 1}"><queryTableFields count="{len(cols)}">{fields}'
                     f'</queryTableFields></queryTableRefresh></queryTable>').encode()
        order.append(qt)
        ct_add.append(f'<Override PartName="/{qt}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/>')
        trel = tf.replace("xl/tables/", "xl/tables/_rels/") + ".rels"
        files[trel] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">'
                       f'<Relationship Id="rId1" Type="{REL}/queryTable" Target="../queryTables/queryTable{n}.xml"/></Relationships>').encode()
        order.append(trel)
        # tên vùng ExternalData_n
        sheet_file = None
        for f, data in files.items():
            if f.startswith("xl/worksheets/_rels/") and tf.split("/")[-1].encode() in data:
                sheet_file = "xl/worksheets/" + f.split("/")[-1][:-5]
        idx, sname = next((i, s) for i, (s, rid) in enumerate(sheets) if rid_target[rid].lstrip("/").replace("xl/", "") == sheet_file.replace("xl/", ""))
        a, b = ref.split(":")
        abs_ref = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", a) + ":" + re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", b)
        defined.append(f'<definedName name="ExternalData_{n}" localSheetId="{idx}" hidden="1">{sname}!{abs_ref}</definedName>')
        conns.append(f'<connection id="{n}" keepAlive="1" name="Query - {q}" description="Connection to the \'{q}\' query in the workbook." '
                     f'type="5" refreshedVersion="8" background="1" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;'
                     f'Data Source=$Workbook$;Location={q};Extended Properties=&quot;&quot;" command="SELECT * FROM [{q}]"/></connection>')

    files["xl/connections.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="{MAIN}">'
                                   + "".join(conns) + "</connections>").encode()
    order.append("xl/connections.xml")
    ct_add.append('<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/>')
    rel_add.append(f'<Relationship Id="rIdConn1" Type="{REL}/connections" Target="connections.xml"/>')

    # --- DataMashup
    item, _ = data_mashup(queries, loaded, counts)
    files["customXml/item1.xml"] = item
    files["customXml/itemProps1.xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="'
                                         + _guid() + '" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
                                         '<ds:schemaRefs/></ds:datastoreItem>').encode()
    files["customXml/_rels/item1.xml.rels"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">'
                                               f'<Relationship Id="rId1" Type="{REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>').encode()
    order += ["customXml/item1.xml", "customXml/itemProps1.xml", "customXml/_rels/item1.xml.rels"]
    ct_add.append('<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>')
    rel_add.append(f'<Relationship Id="rIdCx1" Type="{REL}/customXml" Target="../customXml/item1.xml"/>')

    # --- công thức mảng động
    if "xl/metadata.xml" not in files:
        files["xl/metadata.xml"] = META_DA.encode()
        order.append("xl/metadata.xml")
        ct_add.append('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>')
        rel_add.append(f'<Relationship Id="rIdMeta1" Type="{REL}/sheetMetadata" Target="metadata.xml"/>')
    for f in list(files):
        if f.startswith("xl/worksheets/sheet") and f.endswith(".xml"):
            s = files[f].decode()
            s = re.sub(r'<c ([^>]*?)>(<f t="array")', lambda m: m.group(0) if "cm=" in m.group(1) else f'<c {m.group(1)} cm="1">{m.group(2)}', s)
            files[f] = s.encode()

    files["[Content_Types].xml"] = files["[Content_Types].xml"].decode().replace("</Types>", "".join(ct_add) + "</Types>").encode()
    files["xl/_rels/workbook.xml.rels"] = rels.replace("</Relationships>", "".join(rel_add) + "</Relationships>").encode()
    wbx = re.sub(r"<definedNames\s*/>", "", wbx)
    if "<definedNames>" in wbx:
        wbx = wbx.replace("</definedNames>", "".join(defined) + "</definedNames>")
    else:
        wbx = wbx.replace("</sheets>", "</sheets><definedNames>" + "".join(defined) + "</definedNames>")
    files["xl/workbook.xml"] = wbx.encode()

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for f in dict.fromkeys(order):
            z.writestr(f, files[f])
