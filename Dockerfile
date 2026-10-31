FROM python:3.13-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py docx_reader.py docx_converter.py xml_builder.py xml_validator.py schema.xsd ./
ENTRYPOINT ["python", "app.py"]
