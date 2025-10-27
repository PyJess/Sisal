from langchain_openai import ChatOpenAI
from dotenv import load_dotenv
import os

load_dotenv()

async def a_invoke_model(msgs, schema, model="gpt-4o-mini"):
        """Invoke the LLM model"""

        gpt = ChatOpenAI(model=model, temperature=0.1).with_structured_output(schema=schema, strict=True)


        result = await gpt.ainvoke(msgs)
    
        # Estrai i token usage
        if isinstance(result, dict) and 'raw' in result:
            usage = result['raw'].response_metadata.get('token_usage', {})
            print(f"\n📊 Token Usage:")
            print(f"  Input tokens:  {usage.get('prompt_tokens', 0)}")
            print(f"  Output tokens: {usage.get('completion_tokens', 0)}")
            print(f"  Total tokens:  {usage.get('total_tokens', 0)}")
            
            # Restituisci solo i dati parsati
            return result['parsed']
        
        return result
        #return await gpt.ainvoke(msgs)

