from langchain_openai import ChatOpenAI
from dotenv import load_dotenv

load_dotenv()

async def a_invoke_model(msgs, schema, model="gpt-4.1"):
        """Invoke the LLM model"""

        gpt = ChatOpenAI(model=model, temperature=0.1).with_structured_output(schema=schema, strict=True)
        return await gpt.ainvoke(msgs)

