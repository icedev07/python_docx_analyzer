# SAMBANOVA_API_URL = "https://fast-api.snova.ai/v1/chat/completions"
# SAMBANOVA_API_KEY = "c3VzbG92d2ViaGVyb0BnbWFpbC5jb206eU9RbEdyb01oUUlQa19lWA"

from openai import OpenAI
client = OpenAI(
    base_url='https://fast-api.snova.ai/v1',
    api_key='c3VzbG92d2ViaGVyb0BnbWFpbC5jb206eU9RbEdyb01oUUlQa19lWA==',
)

# model = 'llama3-405b'
model = 'Meta-Llama-3.1-8B-Instruct'
prompt = 'Tell me a joke about artificial intelligence.'

completion = client.chat.completions.create(
    model=model,
    messages=[
        {
        'role': 'user',
        'content': prompt,
        }
    ],
    stream=True,
)
response = ''
for chunk in completion:
  response += chunk.choices[0].delta.content or ''

print(response) 
