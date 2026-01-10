import os

import google.genai
from opik import configure 
from opik.integrations.genai import track_genai 

configure() 

# os.environ["GEMINI_API_KEY"] = "your-api-key-here"

client = google.genai.Client()
gemini_client = track_genai(client) 
response = gemini_client.models.generate_content(
    model="gemini-2.0-flash-001", contents="Write a haiku about AI engineering."
)
print(response.text)
