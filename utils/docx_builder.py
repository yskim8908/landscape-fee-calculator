from io import BytesIO
from docxtpl import DocxTemplate


def build_directive_docx(template_path: str, context: dict) -> BytesIO:
    """
    과업지시서(.docx) 생성 함수
    
    Args:
        template_path: 템플릿 파일 경로
        context: 템플릿에 채울 데이터 딕셔너리
    
    Returns:
        BytesIO: 생성된 Word 파일
    """
    tpl = DocxTemplate(template_path)
    tpl.render(context)
    buf = BytesIO()
    tpl.save(buf)
    buf.seek(0)
    return buf

