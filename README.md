# ManyPython
Many Python code

# curl test

### GET /
```bash
curl -v http://127.0.0.1:8000/
```
sample response
```json
{"message":"Welcome to the FastAPI application: Hi, master"}
```

### GET response
```bash
curl http://127.0.0.1:8000/response/
```
sample response
```json
{"message":"No response available."}
```

### POST clearPromptResponse
```bash
curl -X POST http://127.0.0.1:8000/clearPromptResponse/
```
sample response
```json
{"message":"Prompt, response cleared successfully!"}
```

### POST prompt
```bash
curl -X POST -H "Content-Type: application/json" -d "{\"prompt\":\"Open google.com website\"}" http://127.0.0.1:8000/prompt/
```
sample response
```bash
{"message":"Prompt created successfully!","prompt":"Open wikipedia"}
```