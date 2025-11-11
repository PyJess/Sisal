from langchain_openai import ChatOpenAI
from dotenv import load_dotenv
import os

load_dotenv()

# async def a_invoke_model(msgs, schema, model="gpt-4o-mini"):
#         """Invoke the LLM model"""

#         gpt = ChatOpenAI(model=model, temperature=0.1).with_structured_output(schema=schema, strict=True)


#         result = await gpt.ainvoke(msgs)
    
#         # Estrai i token usage
#         if isinstance(result, dict) and 'raw' in result:
#             usage = result['raw'].response_metadata.get('token_usage', {})
#             print(f"\n📊 Token Usage:")
#             print(f"  Input tokens:  {usage.get('prompt_tokens', 0)}")
#             print(f"  Output tokens: {usage.get('completion_tokens', 0)}")
#             print(f"  Total tokens:  {usage.get('total_tokens', 0)}")
            
#             # Restituisci solo i dati parsati
#             return result['parsed']
        
#         return result
        #return await gpt.ainvoke(msgs)



async def a_invoke_model_without_schema(msgs, model="gpt-4o-mini"):
        """Invoke the LLM model"""

        gpt = ChatOpenAI(model=model, temperature=0.1)

        result = await gpt.ainvoke(msgs)
        return result




from langchain_openai import ChatOpenAI
import os
from dotenv import load_dotenv

from pydantic import BaseModel, Field
from typing import List, Optional, Literal

# class TestStep(BaseModel):
#     Step: int = Field(description="Sequential number of the step")
#     Step_Description: str = Field(alias="Step Description", description="Detailed description of the action")
#     Expected_Result: str = Field(alias="Expected Result", description="Expected outcome after executing this step")

#     class Config:
#         populate_by_name = True

# class TestCase(BaseModel):
#     Title: str = Field(description="Concise descriptive title")
#     ID: str = Field(description="Unique identifier")
#     Steps: List[TestStep] = Field(description="List of ordered steps")
#     Test_Group: str = Field(alias="Test Group", description="Device category")
#     Channel: str = Field(description="Channel or interface")
#     Device: str = Field(description="Hardware device")
#     Priority: Literal["High", "Medium", "Low"] = Field(description="Execution priority")
#     Test_Stage: str = Field(alias="Test Stage", description="Testing environment")
#     Reference_System: str = Field(alias="Reference System", description="Functional area")
#     Preconditions: str = Field(description="Required conditions")
#     Execution_Mode: Literal["Manual", "Automated"] = Field(alias="Execution Mode", description="Execution mode")
#     Functionality: str = Field(description="Specific functionality")
#     Test_Type: Literal["functional", "automation", "ux_ui", "content", "vapt", "performance", "seo"] = Field(alias="Test Type", description="Main category")
#     No_Regression_Test: bool = Field(alias="No Regression Test", description="Include in regression")
#     Automation: bool = Field(description="Consider for automation")
#     Dataset: str = Field(description="Required data")
#     Expected_Result: str = Field(alias="Expected Result", description="Overall expected result")
#     Country: str = Field(description="Geographical indication")
#     Type: str = Field(description="Polarion element type")
#     Partial_Coverage_Description: Optional[str] = Field(alias="Partial Coverage Description", default=None, description="Partial coverage description")
#     polarion: str = Field(alias="_polarion", description="Polarion ID(s)")  # ⬅ CORRETTO!
from langchain_openai import ChatOpenAI
import os
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')

def get_token():
    import requests
    url = os.getenv('BASE_URL_TOKEN')
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'api://3bb3eccb-0787-4526-811e-ec3dab677121/.default',
        'grant_type': 'client_credentials'
    }
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    response = requests.post(url, data=payload, headers=headers)
    if response.status_code == 200:
        return response.json().get('access_token')
    else:
        print(f"Error Token: {response.status_code} - {response.text}")
        return None

class LLMClient:
    def __init__(
            self,
            model_name: str = "gpt-4o-mini",
            temperature=0,
    ):
        self.model_name = model_name
        token = get_token()
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Ocp-Apim-Subscription-Key": os.getenv("OCP_APIM_SUBSCRIPTION_KEY")
        }
        self.openai_api_base = os.getenv("OPENAI_BASE_URL")
        self.temperature = temperature
        try:
            self._client = ChatOpenAI(
                model=self.model_name,
                default_headers=self.headers,
                openai_api_base=self.openai_api_base,
                temperature=self.temperature,
                api_key="useless",
            )
        except Exception as e:
            print(f"[LLMClient] initialization error: {e}")
            self._client = None
    
    async def a_invoke_model(self, msgs, schema):
        """Invoke the LLM model with structured output"""
        if self._client is None:
            raise ValueError("LLM Client non inizializzato correttamente")
        
        # Configura structured output
        llm_with_structure = self._client.with_structured_output(
            schema=schema, 
            strict=True
        )
        
        # Invoca il modello
        result = await llm_with_structure.ainvoke(msgs)
        
        # Estrai i token usage se disponibili
        if hasattr(result, '__dict__'):
            usage_info = getattr(result, 'usage_metadata', None) or \
                        getattr(result, 'response_metadata', {}).get('token_usage', {})
            
            if usage_info:
                print(f"\n📊 Token Usage:")
                print(f"  Input tokens:  {usage_info.get('input_tokens', usage_info.get('prompt_tokens', 0))}")
                print(f"  Output tokens: {usage_info.get('output_tokens', usage_info.get('completion_tokens', 0))}")
                print(f"  Total tokens:  {usage_info.get('total_tokens', 0)}")
        
        return result
    


from langchain_openai import OpenAIEmbeddings

class EmbeddingsClient:
    def __init__(self, model_name: str = "text-embedding-3-large"):
        self.model_name = model_name
        token = get_token()
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Ocp-Apim-Subscription-Key": os.getenv("OCP_APIM_SUBSCRIPTION_KEY")
        }
        self.openai_api_base = os.getenv("OPENAI_BASE_URL")
        
        try:
            self._client = OpenAIEmbeddings(
                model=self.model_name,
                default_headers=self.headers,
                openai_api_base=self.openai_api_base,
                api_key="useless",
            )
        except Exception as e:
            print(f"[EmbeddingsClient] initialization error: {e}")
            self._client = None
    
    def get_embeddings(self):
        """Restituisce il client embeddings configurato"""
        if self._client is None:
            raise ValueError("Embeddings Client non inizializzato correttamente")
        return self._client

