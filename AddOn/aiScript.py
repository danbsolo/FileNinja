from openai import OpenAI
from pydantic import BaseModel
client = OpenAI()

class CorrectedFileNames(BaseModel):
    correctedFileNamesList: list[str]

def queryAI(userInput, developerInstructions=None):
    response = client.responses.parse(
        model="gpt-4.1",
        instructions=developerInstructions,
        input=userInput,
        text_format=CorrectedFileNames,
        max_output_tokens=50_000
    )

    return response

# if __name__ == "__main__":
#     userInput = input("Say something to the AI model: ")
#     print()
#     print(queryAI(userInput).output_text)
