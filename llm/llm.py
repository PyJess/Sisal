from langchain_openai import ChatOpenAI


async def a_invoke_model(msgs, schema):
        """Invoke the LLM model"""

        gpt = ChatOpenAI(model="gpt-4.1", temperature=0.1).with_structured_output(schema=schema, strict=True)
        return await gpt.ainvoke(msgs)

