"""
FastAPI 应用：HTML 转 Word 文档服务
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel
import tempfile
import os
from html_to_word import html_to_word

app = FastAPI(
    title="HTML to Word Converter",
    description="将 HTML 格式的简历转换为 Word 文档",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class HTMLContent(BaseModel):
    """HTML 内容请求模型"""
    html_content: str


@app.get("/")
async def root():
    """根路径，返回 API 信息"""
    return {
        "message": "HTML to Word Converter API",
        "version": "1.0.0",
    }


@app.post("/html2word")
async def html2word(html_content: HTMLContent):
    try:
        # 验证 HTML 内容
        if not html_content.html_content or not html_content.html_content.strip():
            raise HTTPException(status_code=400, detail="HTML 内容不能为空")

        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
            tmp_path = tmp_file.name

        try:
            # 转换 HTML 为 Word
            doc = html_to_word(html_content.html_content, output_path=tmp_path)

            # 返回文件（使用 Response 以便在下载后删除临时文件）
            def remove_file():
                if os.path.exists(tmp_path):
                    os.unlink(tmp_path)

            # 读取文件内容
            with open(tmp_path, 'rb') as f:
                file_content = f.read()

            # 删除临时文件
            remove_file()

            return Response(
                content=file_content,
                media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                headers={"Content-Disposition": "attachment; filename=output.docx"}
            )
        except Exception as e:
            # 如果出错，删除临时文件
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
            raise HTTPException(status_code=500, detail=f"转换失败: {str(e)}")

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"服务器错误: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
