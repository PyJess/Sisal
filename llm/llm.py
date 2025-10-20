from langchain_openai import ChatOpenAI
from dotenv import load_dotenv
import os

load_dotenv()


async def a_invoke_model(msgs, schema):
        """Invoke the LLM model"""

        gpt = ChatOpenAI(model="gpt-5-mini", temperature=0.1).with_structured_output(schema=schema, strict=True)
        return await gpt.ainvoke(msgs)

