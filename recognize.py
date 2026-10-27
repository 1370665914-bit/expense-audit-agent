# -*- coding: utf-8 -*-
"""
凭证识别模块 —— 「AI 填单」的眼睛
================================
职责:发票/截图/行程单 → 文字 → 结构化字段(发票号/日期/金额/商户/费用类型)

识别手段(全部离线,不联网):
  · 图片(jpg/png/heic):macOS 系统自带的 Vision 框架 OCR
  · PDF(电子发票/滴滴行程单):pypdf 提取内嵌文字;
    扫描版 PDF 提不到文字时,用 Quartz 把首页渲染成图片再 OCR
"""
import re
import tempfile
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".heic", ".bmp", ".tiff", ".webp"}


# ---------------------------------------------------------------------------
# 文字提取
# ---------------------------------------------------------------------------

def ocr_image(path):
    """用 macOS Vision 框架做离线 OCR,返回按行拼接的文本"""
    import Vision
    from Foundation import NSURL
    url = NSURL.fileURLWithPath_(str(path))
    handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(url, None)
    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLanguages_(["zh-Hans", "en-US"])
    req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    req.setUsesLanguageCorrection_(True)
    ok, err = handler.performRequests_error_([req], None)
    if not ok:
        raise RuntimeError("OCR 失败:%s" % err)
    lines = []
    for obs in (req.results() or []):
        cands = obs.topCandidates_(1)
        if cands and len(cands):
            lines.append(str(cands[0].string()))
    return "\n".join(lines)


def pdf_page_to_png(path, out_png, scale=2.5):
    """把 PDF 第一页渲染成 PNG(用于扫描版 PDF 的 OCR 兜底)"""
    import Quartz
    from Foundation import NSURL
    doc = Quartz.CGPDFDocumentCreateWithURL(NSURL.fileURLWithPath_(str(path)))
    if doc is None or Quartz.CGPDFDocumentGetNumberOfPages(doc) == 0:
        raise RuntimeError("PDF 无法打开")
    page = Quartz.CGPDFDocumentGetPage(doc, 1)
    rect = Quartz.CGPDFPageGetBoxRect(page, Quartz.kCGPDFMediaBox)
    w, h = int(rect.size.width * scale), int(rect.size.height * scale)
    cs = Quartz.CGColorSpaceCreateDeviceRGB()
    ctx = Quartz.CGBitmapContextCreate(None, w, h, 8, 0, cs, Quartz.kCGImageAlphaPremultipliedLast)
    Quartz.CGContextSetRGBFillColor(ctx, 1, 1, 1, 1)
    Quartz.CGContextFillRect(ctx, Quartz.CGRectMake(0, 0, w, h))
    Quartz.CGContextScaleCTM(ctx, scale, scale)
    Quartz.CGContextDrawPDFPage(ctx, page)
    img = Quartz.CGBitmapContextCreateImage(ctx)
    dest = Quartz.CGImageDestinationCreateWithURL(
        Quartz.CFURLCreateWithFileSystemPath(None, str(out_png), Quartz.kCFURLPOSIXPathStyle, False),
        "public.png", 1, None)
    Quartz.CGImageDestinationAddImage(dest, img, None)
    Quartz.CGImageDestinationFinalize(dest)


def extract_text(path):
    """任意凭证文件 → (文本, 识别方式)"""
    p = Path(path)
    ext = p.suffix.lower()
    if ext in IMAGE_EXTS:
        return ocr_image(p), "图片OCR"
    if ext == ".pdf":
        from pypdf import PdfReader
        try:
            text = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages)
        except Exception:
            text = ""
        if len(text.strip()) >= 20:
            return text, "PDF文本"
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            pdf_page_to_png(p, tmp.name)
            return ocr_image(tmp.name), "PDF转图OCR"
    raise ValueError("暂不支持的文件类型:%s(支持 jpg/png/heic/pdf)" % ext)


# ---------------------------------------------------------------------------
# 字段解析(正则 + 关键词,全部可离线)
# ---------------------------------------------------------------------------

TYPE_KEYWORDS = [
    ("餐饮费", ["餐饮", "餐费", "饮食", "美食", "用餐", "食品"]),
    ("交通费", ["出租", "客运", "运输服务", "滴滴", "行程单", "打车", "网约车", "高德", "曹操", "地铁", "公交"]),
    ("差旅费", ["住宿", "酒店", "宾馆", "机票", "火车票", "高铁", "航空"]),
    ("办公用品", ["办公", "文具", "耗材", "纸张", "打印", "墨盒"]),
    ("招待费", ["招待"]),
]


def _find_amount(text):
    """优先取「价税合计/总金额/实付」附近的小写金额,其次取全文最大的 ¥ 金额"""
    m = re.search(r"(?:价税合计|合计金额|总金额|实付款?|应付金额|合\s*计)[^0-9¥￥]{0,20}[¥￥]?\s*([0-9,]+\.\d{1,2})", text)
    if m:
        return float(m.group(1).replace(",", ""))
    amounts = [float(a.replace(",", "")) for a in re.findall(r"[¥￥]\s*([0-9,]+\.\d{1,2})", text)]
    if amounts:
        return max(amounts)
    m = re.search(r"([0-9,]+\.\d{2})\s*元", text)
    return float(m.group(1).replace(",", "")) if m else None


def _find_date(text):
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return ""


def _find_invoice_no(text):
    m = re.search(r"发票号码\s*[:：]?\s*(\d{8,20})", text)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{20})\b", text)          # 全电发票 20 位号码
    if m:
        return m.group(1)
    m = re.search(r"No\s*[.:：]?\s*(\d{8,12})", text, re.I)
    return m.group(1) if m else ""


def _find_seller(text):
    m = re.search(r"销\s*售\s*方[\s\S]{0,60}?名\s*称\s*[:：]?\s*([^\s:：]{4,30})", text)
    if m:
        return m.group(1)
    for line in text.splitlines():
        line = line.strip()
        if 4 <= len(line) <= 30 and re.search(r"(有限公司|有限责任公司|商店|饭店|酒店|餐厅|中心|商行)$", line):
            return re.sub(r"^.*?[:：]", "", line)
    return ""


def guess_type(text):
    for ftype, words in TYPE_KEYWORDS:
        if any(w in text for w in words):
            return ftype
    return ""


def parse_fields(text):
    """凭证文字 → 报销草稿字段;识别不出的留空,由报销人补全(AI 只生成草稿)"""
    amount = _find_amount(text)
    fields = {
        "发票号": _find_invoice_no(text),
        "开票日期": _find_date(text),
        "金额": amount,
        "商户": _find_seller(text),
        "费用类型": guess_type(text),
    }
    missing = [k for k, v in fields.items() if v in ("", None)]
    fields["识别备注"] = ("已识别全部关键字段" if not missing
                          else "未识别:%s(请人工补全)" % "、".join(missing))
    return fields


def recognize(path):
    """一站式:文件 → 字段 dict(含 识别方式/识别备注/原始文本)"""
    text, method = extract_text(path)
    fields = parse_fields(text)
    fields["识别方式"] = method
    fields["原始文本"] = text
    return fields
