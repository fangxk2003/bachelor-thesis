import copy
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


BASE = Path(__file__).resolve().parent
SRC = BASE / "mid-term-presentation.pptx"
OUT = BASE / "final-defense-zip-edited.pptx"
WORK = Path("/private/tmp/midterm_zip_edit_work")

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for k, v in NS.items():
    ET.register_namespace(k, v)


def q(ns, tag):
    return f"{{{NS[ns]}}}{tag}"


KEEP = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 16, 17, 18, 19, 22]


def unzip():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    with zipfile.ZipFile(SRC, "r") as z:
        z.extractall(WORK)


def zipdir():
    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(WORK.rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(WORK).as_posix())


def list_slides_only():
    pres = WORK / "ppt" / "presentation.xml"
    root = ET.parse(pres)
    pr = root.getroot()
    sld_id_lst = pr.find("p:sldIdLst", NS)
    old = list(sld_id_lst)
    rid_by_slide = {}
    for idx, node in enumerate(old, start=1):
        rid_by_slide[idx] = node.attrib[q("r", "id")]
    for node in old:
        sld_id_lst.remove(node)
    next_id = 256
    for slide_no in KEEP:
        node = ET.Element(q("p", "sldId"), {"id": str(next_id), q("r", "id"): rid_by_slide[slide_no]})
        sld_id_lst.append(node)
        next_id += 1
    root.write(pres, encoding="UTF-8", xml_declaration=True)


def update_shape_text(slide_no, shape_idx, lines, size=None):
    path = WORK / "ppt" / "slides" / f"slide{slide_no}.xml"
    tree = ET.parse(path)
    root = tree.getroot()
    shapes = [sp for sp in root.findall(".//p:sp", NS) if sp.find(".//p:txBody", NS) is not None]
    if shape_idx >= len(shapes):
        tree.write(path, encoding="UTF-8", xml_declaration=True)
        return
    sp = shapes[shape_idx]
    tx = sp.find("p:txBody", NS)
    body_pr = tx.find("a:bodyPr", NS)
    lst_style = tx.find("a:lstStyle", NS)
    first_rpr = sp.find(".//a:rPr", NS)
    rpr_template = copy.deepcopy(first_rpr) if first_rpr is not None else None
    for child in list(tx):
        if child.tag == q("a", "p"):
            tx.remove(child)
    for line in lines if isinstance(lines, list) else str(lines).split("\n"):
        p = ET.Element(q("a", "p"))
        r = ET.SubElement(p, q("a", "r"))
        if rpr_template is not None:
            r.append(copy.deepcopy(rpr_template))
        else:
            ET.SubElement(r, q("a", "rPr"), {"lang": "zh-CN"})
        rpr = r.find("a:rPr", NS)
        if size is not None and rpr is not None:
            rpr.set("sz", str(int(size * 100)))
        t = ET.SubElement(r, q("a", "t"))
        t.text = line
        tx.append(p)
    # keep bodyPr/lstStyle at the beginning if PowerPoint moved them
    ordered = []
    if body_pr is not None:
        try:
            tx.remove(body_pr)
        except ValueError:
            pass
        ordered.append(body_pr)
    if lst_style is not None:
        try:
            tx.remove(lst_style)
        except ValueError:
            pass
        ordered.append(lst_style)
    rest = list(tx)
    for c in rest:
        tx.remove(c)
    for c in ordered + rest:
        tx.append(c)
    tree.write(path, encoding="UTF-8", xml_declaration=True)


def blank_shape_text(slide_no, shape_idxs):
    for idx in shape_idxs:
        update_shape_text(slide_no, idx, [""])


def remove_slide_pictures(slide_no):
    path = WORK / "ppt" / "slides" / f"slide{slide_no}.xml"
    tree = ET.parse(path)
    root = tree.getroot()
    for parent in root.iter():
        for child in list(parent):
            if child.tag == q("p", "pic"):
                parent.remove(child)
    tree.write(path, encoding="UTF-8", xml_declaration=True)


def edit_texts():
    # slide 1 cover
    update_shape_text(1, 0, ["基于深度学习的膝关节 ACL/MLKI 撕裂后", "骨与软骨分割及 BML 量化分析研究"], size=30)
    update_shape_text(1, 1, ["方行楷  522031910108", "导师：程荣山", "生物医学工程学院 · 2026 年 5 月"], size=18)

    # slide 2 contents
    update_shape_text(2, 0, ["1"])
    update_shape_text(2, 1, ["2"])
    update_shape_text(2, 2, ["5"])
    update_shape_text(2, 3, ["3"])
    update_shape_text(2, 4, ["选题背景与研究问题"])
    update_shape_text(2, 5, ["数据集与技术路线"])
    update_shape_text(2, 6, ["量化指标与统计结果"])
    update_shape_text(2, 7, ["骨与软骨分割方法"])
    update_shape_text(2, 8, ["2"])

    # section slides
    update_shape_text(3, 0, ["选题背景"])
    update_shape_text(3, 1, ["3"])
    update_shape_text(6, 0, ["研究方法"])
    update_shape_text(6, 1, ["6"])
    update_shape_text(8, 0, ["核心工作与成果"])
    update_shape_text(8, 1, ["8"])

    # slide 4 background
    update_shape_text(4, 1, ["选题背景"])
    update_shape_text(4, 2, ["ACL 与 MLKI 损伤后的 MRI 量化需求"])
    update_shape_text(4, 3, [
        "ACL 撕裂常见于运动损伤，MLKI 多与高能创伤相关",
        "两类损伤均会增加创伤后骨关节炎（PTOA）风险",
        "MRI 可同时观察骨、软骨、关节积液与 BML 相关高信号",
    ])
    update_shape_text(4, 7, ["研究痛点"])
    update_shape_text(4, 8, ["三维骨与软骨人工标注成本高；临床 FSE 与公开 DESS 存在跨序列域偏移。"])
    update_shape_text(4, 9, ["BML 边界弥散，缺少大规模精细病灶金标准；需定位为候选高信号体积。"])
    update_shape_text(4, 10, ["研究目标：建立分割与量化流程，比较 ACL 与 MLKI 的影像定量差异。"])
    update_shape_text(4, 11, ["4"])
    update_shape_text(4, 12, ["核心指标：软骨厚度、胫股关节影像近接触面积、BML 候选高信号体积"])

    # slide 5 dataset/goal
    update_shape_text(5, 0, ["数据集与研究目标"])
    update_shape_text(5, 1, [
        "源域：OAI-ZIB，507 例 3D DESS，含骨与软骨标签",
        "目标域：665 例临床 FSE，ACL 500 例，MLKI 165 例",
        "人工测试集：30 例 FSE，约 350 张标注切片；150 张用于定量评价",
        "临床统计：质控后 579 例，ACL 430 例，MLKI 149 例",
    ])
    update_shape_text(5, 2, ["研究目的：在少目标域标注依赖下完成骨与软骨分割，并提取软骨厚度、影像近接触面积和 BML 候选高信号体积，用于 ACL 与 MLKI 的组间比较。"])
    update_shape_text(5, 3, ["5"])
    update_shape_text(5, 4, ["OAI-ZIB\n(DESS)"])
    update_shape_text(5, 5, ["临床 ACL/MLKI\n(FSE)"])
    update_shape_text(5, 6, ["伦理号：2025-KY-295 (K)"])
    update_shape_text(5, 7, ["域迁移"])

    # slide 7 route, keep original flowchart shapes but update key texts
    update_shape_text(7, 0, ["总体技术路线"])
    update_shape_text(7, 1, [
        "完成临床 FSE 与 OAI-ZIB 数据预处理",
        "完成 30 例 / 约 350 张 FSE 切片人工标注",
        "开发改进型 DANN 进行骨与软骨分割",
        "建立软骨厚度、近接触面积、BML 候选体积量化流程",
        "对 ACL 430 例与 MLKI 149 例进行统计比较",
    ])
    update_shape_text(7, 2, ["技术路线图"])
    update_shape_text(7, 3, ["7"])
    update_shape_text(7, 4, ["源域：OAI-ZIB\n507 例 DESS"])
    update_shape_text(7, 7, ["改进型 DANN\n模型训练"])
    update_shape_text(7, 9, ["目标域：ACL/MLKI\n665 例 FSE"])
    update_shape_text(7, 17, ["统计学\n分析"])
    update_shape_text(7, 18, ["分割结果"])
    update_shape_text(7, 19, ["量化结果"])
    update_shape_text(7, 20, ["ACL/MLKI\n组间比较"])

    # slide 9 data
    update_shape_text(9, 0, ["核心工作一：数据集构建与标准化处理"])
    update_shape_text(9, 1, ["目标域测试标注：30 例 FSE，约 350 张切片"])
    update_shape_text(9, 2, ["人工标注用于模型测试，不参与 DANN 训练"])
    update_shape_text(9, 3, ["预处理：DICOM 脱敏、NIfTI 体积化、N4 校正、统一 spacing、三维 Z-score"])

    # slide 11 model
    update_shape_text(11, 0, ["核心工作二：改进型 DANN 分割模型"])
    update_shape_text(11, 1, ["10"])
    update_shape_text(11, 2, ["U-Net 式解码器"])
    update_shape_text(11, 5, ["ASPP 多尺度上下文模块"])
    update_shape_text(11, 7, ["ResNet34 编码器"])
    update_shape_text(11, 8, ["实例归一化减弱 MRI 灰度风格差异；域判别器 + 梯度反转层实现 DESS/FSE 特征级对齐；动态冻结缓解对抗训练失衡。"])

    # slide 16 segmentation result
    update_shape_text(16, 0, ["核心成果一：骨与软骨分割结果"])
    update_shape_text(16, 1, ["11"])
    update_shape_text(16, 2, ["改进型 DANN 在目标域 FSE 人工测试集上的分割结果"])
    update_shape_text(16, 3, [
        "股骨 DSC：0.9557 ± 0.0243，ASD：1.039 mm",
        "胫骨 DSC：0.9678 ± 0.0231，ASD：0.547 mm",
        "股骨软骨 DSC：0.7748 ± 0.0633，ASD：0.961 mm",
        "胫骨软骨 DSC：0.7530 ± 0.1505，ASD：0.678 mm",
    ])
    update_shape_text(16, 4, ["说明：当前结果主要验证最终改进型 DANN 的可行性；Source-only、标准 DANN 和模块消融仍需后续补充。"])
    update_shape_text(16, 5, ["骨与软骨分割结果为后续厚度、近接触面积和 BML 候选体积量化提供基础。"])

    # slide 17 metric system
    update_shape_text(17, 0, ["核心工作三：多维量化评估体系"])
    update_shape_text(17, 1, ["12"])
    update_shape_text(17, 2, ["软骨厚度"])
    update_shape_text(17, 3, ["三维 EDT"])
    update_shape_text(17, 4, ["厚度热力图"])
    update_shape_text(17, 5, ["影像近接触面积"])
    update_shape_text(17, 6, ["BML 候选体积"])
    update_shape_text(17, 7, ["Marching Cubes + kd-tree\n计算 <1 mm 近接触面片"])
    update_shape_text(17, 8, ["BML 定位为候选高信号区域的自动检测、粗分割和体积量化"])
    update_shape_text(17, 9, ["量化指标进入 ACL 与 MLKI 统计比较"])

    # slide 18 thickness
    update_shape_text(18, 0, ["核心工作三：软骨厚度与近接触面积"])
    update_shape_text(18, 1, ["13"])
    update_shape_text(18, 2, ["软骨厚度热力图与胫股关节影像近接触面积可视化"])

    # slide 19 bml/stat result
    remove_slide_pictures(19)
    update_shape_text(19, 0, ["核心成果二：ACL 与 MLKI 统计比较"])
    update_shape_text(19, 1, ["14"])
    update_shape_text(19, 2, ["统计队列：质控后 579 例，ACL 430 例，MLKI 149 例"])
    update_shape_text(19, 3, [
        "影像近接触面积：ACL 中位数 269.33 mm²，MLKI 41.13 mm²；校正后 MLKI 较 ACL 低 110.78 mm²，p=8.53e-9",
        "BML 候选体积：ACL 中位数 3397.55 mm³，MLKI 5214.46 mm³；校正后 MLKI 较 ACL 高 1326.82 mm³，p=5.17e-7",
    ], size=20)

    # slide 22 final
    update_shape_text(22, 0, ["恳请各位老师批评指正！"], size=36)


def main():
    unzip()
    list_slides_only()
    edit_texts()
    zipdir()
    print(OUT)


if __name__ == "__main__":
    main()
