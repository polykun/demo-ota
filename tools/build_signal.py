from __future__ import annotations

import hashlib
import urllib.request
import zipfile
from pathlib import Path
from typing import Final

from lxml import etree

SOURCE_URL: Final[str] = "https://www.jfc.go.jp/n/service/xls/signal_sheet_220401.xlsm"
SOURCE: Final[Path] = Path("signal_source.xlsm")
OUTPUT: Final[Path] = Path("signal_piquant.xlsm")

MAIN_NS: Final[str] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS: Final[str] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS: Final[str] = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS: Final[str] = "http://www.w3.org/XML/1998/namespace"


def qname(tag: str) -> str:
	return f"{{{MAIN_NS}}}{tag}"


def sha256_bytes(data: bytes) -> str:
	return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
	hash_obj = hashlib.sha256()
	with path.open("rb") as handle:
		for chunk in iter(lambda: handle.read(1024 * 1024), b""):
			hash_obj.update(chunk)
	return hash_obj.hexdigest()


def get_sheet_target(workbook_xml: bytes, rels_xml: bytes, sheet_name: str) -> str:
	parser = etree.XMLParser(remove_blank_text=False)
	workbook = etree.fromstring(workbook_xml, parser)
	rels = etree.fromstring(rels_xml, parser)
	rel_map = {
		rel.get("Id"): rel.get("Target")
		for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
	}
	for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
		if sheet.get("name") == sheet_name:
			rel_id = sheet.get(f"{{{REL_NS}}}id")
			if rel_id is None or rel_id not in rel_map:
				raise RuntimeError(f"Relationship not found for sheet: {sheet_name}")
			target = rel_map[rel_id]
			return target if target.startswith("xl/") else f"xl/{target}"
	raise RuntimeError(f"Sheet not found: {sheet_name}")


def find_cell(root: etree._Element, coordinate: str) -> etree._Element:
	cell = root.find(f".//{{{MAIN_NS}}}c[@r='{coordinate}']")
	if cell is None:
		raise RuntimeError(f"Cell not found: {coordinate}")
	return cell


def clear_value_children(cell: etree._Element) -> None:
	for child in list(cell):
		if child.tag in {qname("v"), qname("is"), qname("f")}:
			cell.remove(child)


def set_boolean(root: etree._Element, coordinate: str, value: bool) -> None:
	cell = find_cell(root, coordinate)
	clear_value_children(cell)
	cell.set("t", "b")
	value_element = etree.SubElement(cell, qname("v"))
	value_element.text = "1" if value else "0"


def set_inline_string(root: etree._Element, coordinate: str, value: str) -> None:
	cell = find_cell(root, coordinate)
	clear_value_children(cell)
	cell.set("t", "inlineStr")
	inline_string = etree.SubElement(cell, qname("is"))
	text = etree.SubElement(inline_string, qname("t"))
	text.set(f"{{{XML_NS}}}space", "preserve")
	text.text = value


def set_formula_cache(root: etree._Element, coordinate: str, value: int) -> None:
	cell = find_cell(root, coordinate)
	formula = cell.find(qname("f"))
	if formula is None:
		raise RuntimeError(f"Formula cell not found: {coordinate}")
	cached = cell.find(qname("v"))
	if cached is None:
		cached = etree.SubElement(cell, qname("v"))
	cached.text = str(value)


def modify_workbook_xml(data: bytes) -> bytes:
	parser = etree.XMLParser(remove_blank_text=False)
	root = etree.fromstring(data, parser)
	workbook_view = root.find(f".//{{{MAIN_NS}}}workbookView")
	if workbook_view is not None:
		workbook_view.set("activeTab", "3")
		workbook_view.set("firstSheet", "3")
	calc_pr = root.find(f".//{{{MAIN_NS}}}calcPr")
	if calc_pr is None:
		calc_pr = etree.SubElement(root, qname("calcPr"))
	calc_pr.set("calcMode", "auto")
	calc_pr.set("fullCalcOnLoad", "1")
	calc_pr.set("forceFullCalc", "1")
	calc_pr.set("calcId", "0")
	return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def modify_service_sheet_xml(data: bytes) -> bytes:
	parser = etree.XMLParser(remove_blank_text=False)
	root = etree.fromstring(data, parser)

	question_cells = [
		"G13", "G14", "G15", "G16", "G17", "G18", "G19", "G20",
		"G21", "G22", "G23", "G24", "G25", "G29", "G33", "G35",
	]
	checked = {"G15", "G17", "G19", "G25", "G33"}
	for coordinate in question_cells:
		set_boolean(root, coordinate, coordinate in checked)

	set_inline_string(root, "U14", "AI・Web・ゲームUI・交通データを組み合わせた企画力があり、\n観光DXや出版社IP占い企画まで具体化。\n一方、資金制約により複数事業の同時展開が難しい。")
	set_inline_string(root, "Z14", "営業・販促の専任者がおらず、代表者が開発と兼務している。\n自治体、観光局、出版社、IP保有者及び外部プロデュース会社との連携で\n販路を補強する。")
	set_inline_string(root, "U18", "共通Web基盤とAI技術を再利用し、\n少人数・短期間で高付加価値のシステムを開発できる。\n観光DX700万円案件を実績化する。")
	set_inline_string(root, "Z18", "顧客対応は迅速だが、継続保守及び顧客管理の仕組み化が不足している。\n保守契約、利用ログ及び定期報告の仕組みを整備する。")
	set_inline_string(root, "U22", "企画・開発人材は揃っているが、営業、販促及び顧客管理が\n代表者に集中している。社内の役割分担を明確化し、\n外部パートナーも活用する。")
	set_inline_string(root, "Z22", "既存Web基盤、AI占い基盤、交通データ処理、地図及び経路案内技術を\n保有している。既存資産の再利用により開発工数、外注費及び\n初期投資を抑える。")
	set_inline_string(root, "U29", "受注空白と元金返済再開により一時的に資金不足。\n役員報酬30％削減、代表者債権債務689万円の相殺、分割請求、\n元金据置及び追加運転資金で改善する。")
	set_inline_string(root, "Z29", "AI、競合、市場及び顧客情報を継続的に収集している。\n営業案件、顧客意見、契約状況及び入金予定の一元管理を強化する。")

	set_inline_string(root, "AF14", "生成AIの普及は従来型外注開発には逆風だが、\n自治体DX、多言語観光案内及びAIと出版IPを組み合わせた\nWebサービスには新たな需要が生じている。")
	set_inline_string(root, "AF18", "ゲーム会社のタイトル選別、開発内製化及び外注抑制により、案件数の減少と\n契約判断の長期化が進んでいる。一方、観光DX700万円案件が契約段階まで進み、\n出版社IPを活用する占いサービスの共同企画も具体化している。")
	set_inline_string(root, "AF22", "生成AIを利用した低価格・短納期開発が増え、従来型の受託開発競争が\n激化している。一方、ゲームUI、Web、AI及び交通データを\n一体で扱えることが当社の差別化要因である。")

	set_inline_string(root, "AF25", "☑")
	set_inline_string(root, "AF27", "☑")
	set_inline_string(root, "AF29", "")
	set_inline_string(root, "AF31", "")
	set_inline_string(root, "AF34", "既存借入の元金据置延長に加え、経営環境変化対応資金等による追加運転資金及び\n既存コロナ融資の返済負担軽減策を相談したい。\n観光DX案件の開発・入金までの運転資金に充当する。")
	set_inline_string(root, "AJ37", "株式会社ピカント")

	for coordinate, value in {
		"L13": 1, "M13": 6,
		"L14": 2, "M14": 4,
		"L15": 0, "M15": 8,
		"L16": 2, "M16": 4,
		"Q13": 2, "R13": 4,
		"Q14": 0, "R14": 8,
		"Q15": 3, "R15": 2,
		"Q16": 0, "R16": 8,
	}.items():
		set_formula_cache(root, coordinate, value)

	sheet_view = root.find(f".//{{{MAIN_NS}}}sheetView")
	if sheet_view is not None:
		sheet_view.set("tabSelected", "1")

	return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def main() -> None:
	urllib.request.urlretrieve(SOURCE_URL, SOURCE)
	with zipfile.ZipFile(SOURCE, "r") as source_zip:
		workbook_xml = source_zip.read("xl/workbook.xml")
		rels_xml = source_zip.read("xl/_rels/workbook.xml.rels")
		service_sheet_path = get_sheet_target(workbook_xml, rels_xml, "サービス業編")
		replacements = {
			"xl/workbook.xml": modify_workbook_xml(workbook_xml),
			service_sheet_path: modify_service_sheet_xml(source_zip.read(service_sheet_path)),
		}
		input_entries = source_zip.namelist()
		input_vba_hash = sha256_bytes(source_zip.read("xl/vbaProject.bin"))
		with zipfile.ZipFile(OUTPUT, "w") as output_zip:
			for info in source_zip.infolist():
				data = replacements.get(info.filename, source_zip.read(info.filename))
				output_zip.writestr(info, data)

	with zipfile.ZipFile(OUTPUT, "r") as output_zip:
		bad_entry = output_zip.testzip()
		if bad_entry is not None:
			raise RuntimeError(f"Corrupt ZIP entry: {bad_entry}")
		if output_zip.namelist() != input_entries:
			raise RuntimeError("OOXML entry list changed unexpectedly")
		output_vba_hash = sha256_bytes(output_zip.read("xl/vbaProject.bin"))
		if output_vba_hash != input_vba_hash:
			raise RuntimeError("VBA project changed unexpectedly")

	print(f"exists={OUTPUT.exists()}")
	print(f"size={OUTPUT.stat().st_size}")
	print(f"sha256={sha256_file(OUTPUT)}")
	print(f"zip_entries={len(input_entries)}")
	print("zip_test=OK")
	print(f"vba_sha256={input_vba_hash}")


if __name__ == "__main__":
	main()
