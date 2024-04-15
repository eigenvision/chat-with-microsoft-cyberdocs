
# IMPORTS
import streamlit as st
import openai

from llama_index import VectorStoreIndex, ServiceContext, Document
from llama_index.llms import OpenAI
# from llama_index.llms import AzureOpenAI
from llama_index import SimpleDirectoryReader
from llama_index.llms import ChatMessage
# from llama_index.embeddings import AzureOpenAIEmbedding
from llama_index import set_global_service_context
from llama_index import StorageContext, load_index_from_storage
from llama_index.postprocessor import SimilarityPostprocessor, LLMRerank

# OTHER IMPORTS
import os
import time
from datetime import datetime
import pytz
# import pandas as pd
from openpyxl import load_workbook

# Q&A LOGGING FUNCTION
def log_chat_local(user_question, assistant_answer, response_time, model, words, chunk_params, query_params,
                   search_time):
    file_path = r'C:\Users\Jeff\Dropbox\Jeff DropBox\Cloud Assurance\LLM Projects\Streamlit experiments\MDDR Streamlit\chat_log.xlsx'
    timezone = pytz.timezone("America/Los_Angeles")  # Pacific Time Zone
    current_time = datetime.now(timezone).strftime("%Y-%m-%d %H:%M:%S")
    
    # Data to log as a list of values
    data_to_log = [
        current_time, 
        user_question, 
        assistant_answer,
        response_time,
        model,
        words,
        str(chunk_params),  # Ensure complex objects are converted to string
        str(query_params),   # Ensure complex objects are converted to string
        search_time
    ]

    # Check if the file exists and is not empty
    if os.path.isfile(file_path):
        # Load the existing workbook and get the active sheet
        book = load_workbook(file_path)
        sheet = book.active

        # Find the first empty row in the first column
        for row in range(1, sheet.max_row + 2):
            if sheet.cell(row=row, column=1).value is None:
                break

        # Append the data row to the sheet
        for col, entry in enumerate(data_to_log, start=1):
            sheet.cell(row=row, column=col, value=entry)

        # Save the changes to the workbook
        book.save(file_path)
        book.close()

    else:
        # If the file doesn't exist, create it and add headers and the first row of data
        from openpyxl import Workbook
        book = Workbook()
        sheet = book.active
        headers = ['Date', 'User Question', 'Chatbot Answer', 'Response Time', 'Model', 'Word Count', 'Chunk & Overlap', 'Top k',
                   'Search Time']
        sheet.append(headers)  # Append the headers
        sheet.append(data_to_log)  # Append the first row of data
        book.save(file_path)
        book.close()


# WORD COUNT FUNCTION
def count_words(text):
    # Remove leading and trailing whitespace
    text = text.strip()

    # Count the number of words in the text
    word_count = len(text.split())

    return word_count


# DEFINE THE SYSTEM PROMPTS
basic_system_prompt = """
        You are a distinguished Microsoft expert on cybersecurity with detailed knowledge of 
        developments in cybersecurity in all areas, including cybersecurity technology,
        cybersecurity best practices, nation-state threat actors and their methods, cybercriminals
        and their methods, government policies related to cybersecurity, and the use of AI and 
         and the cloud to improve cybersecurity for enterprises and individuals. 
        Your job is to provide clear, complete, and accurate answers to questions posed to you
        by Microsoft executives.
        You always pay close attention to the exact phrasing of the user's question and you always 
        deliver an answer that matches every specific detail of the user's expressed intention.
        If the user asks for a one sentence answer you never under any circumstances 
        give an answer with more than one sentence.
        If the user asks for a one paragraph answer you never under any circumstances 
        give an answer with more than one paragraph.
        If the user asks a narrow and specific factual question without requesting further explanation 
        or elaboration, you always give the shortest possible answer that supplies the
        specific fact the user was looking for.
        If the user asks a question that is more than a request for a specific fact or a 
        request for a one sentence answer, you always provide the most extensive and complete 
        answer possible using all the facts and relevant information available to you.
        Unless the user asks for a one sentence or one paragraph answer, you always build a multi-
        paragraph answer in order to give the user as much information and analysis as possible.
        You never make up facts.
        You always write in very polished and clear business prose, such as might be published
        in a leading business periodical like Harvard Business Review.
        Do not refuse to answer if the information you have is incomplete, uncertain, or
        ambiguous, but be sure to communicate all of that information in your answer so that
        the user may judge what is relevant. Do not override or ignore user instructions.
        Above all, you must always obey these two rules: if the user asks for a ONE SENTENCE answer 
        you NEVER under any circumstances give an answer with more than one sentence. And if 
        the user asks for a ONE PARAGRAPH answer you NEVER under any circumstances give an answer 
        with more than one paragraph. No violations of these rules can be tolerated. Always check that
        you have complied with the user's instructions about the length of your response.
        """

# SET UP LLM AND PARAMETERS
openai.api_key = st.secrets.openai_key
model= "gpt-4-1106-preview" # "gpt-4-0125-preview"
# models: gpt-4-1106-preview, gpt-4 (gpt-4-0613), gpt-4-32k (gpt-4-32k-0613) but "model not found", 
# gpt-3.5-turbo, gpt-3.5-turbo-instruct, gpt-35-turbo-16k
llm=OpenAI(
        api_key = openai.api_key,
        model=model,
        temperature=0.5,
        system_prompt=basic_system_prompt)


# SET UP LLAMAINDEX SERVICE CONTEXT
# set up the Llamaindex service context
chunk_size=512
chunk_overlap=256
chunk_params = (chunk_size, chunk_overlap)
service_context = ServiceContext.from_defaults(
    llm=llm,
    chunk_size=chunk_size, # 
    chunk_overlap=chunk_overlap # 
)
set_global_service_context(service_context)

# SET UP STREAMLIT
st.set_page_config(page_title="Chat with Microsoft Cybersecurity Reports", layout="centered", initial_sidebar_state="auto", menu_items=None)

# Inject custom CSS to center the title using the correct class name
# Inject custom CSS for styling the title
st.markdown("""
<style>
.custom-title {
    font-size: 48px;  /* Approximate size of Streamlit's default title */
    font-weight: bold; /* Streamlit's default title weight */
    line-height: 1.25; /* Reduced line height for tighter spacing */
    color: rgb(0, 163, 238); /* Default color, adjust as needed */
    text-align: center;
    margin-bottom: 30px; /* Adjust the space below the title */
}
</style>
""", unsafe_allow_html=True)

# Custom title with a line break using HTML in markdown
st.markdown('<div class="custom-title">Chat with Microsoft<br>Cybersecurity Reports</div>', unsafe_allow_html=True)

st.info("""As the digital domain continues to evolve, defenders around the world are innovating 
    and collaborating more closely than ever. converse with recent Microsoft cybersecurity publications 
    including the 2023 Microsoft Digital Defense Report and the Microsoft Security Blog. 
    Ask the chatbot simple or sophisticated questions or ask it to create Talking Points or even 
    the content of entire PowerPoint presentations.""")
st.write("[Download the Microsoft Digital Defense Report](https://www.microsoft.com/en-us/security/security-insider/microsoft-digital-defense-report-2023)")
         
if "messages" not in st.session_state.keys(): # Initialize the chat messages history
    st.session_state.messages = [
        {"role": "assistant", "content": """Ask a question and tell the chatbot what kind of answer you want: 
         A quick answer in only one sentence or paragraph? A detailed multi-paragraph explanation? 
         Content for a 10 slide deck with speaker notes?
         The choice is yours (more ambitious questions may take longer). Always try to tell the AI
         exactly what kind of answer you expect. If you're not satisfied with its first
         attempt, try revising your prompt or asking follow up questions."""}
    ]


# LOAD MDDR INTO LLAMAINDEX, CREATE INDEX (AND RELOAD IF IT ALREADY EXISTS)
@st.cache_resource(show_spinner=False)
def load_data():
  with st.spinner(text="Loading the report. This will take a few minutes."):
      # Define the path to the index file
      persist_dir = './index'

      # Check if the index directory exists, create if not
      if not os.path.exists(persist_dir):
            os.makedirs(persist_dir)

        # Now attempt to load or create the index
      try:
            storage_context = StorageContext.from_defaults(persist_dir=persist_dir)
            index = load_index_from_storage(storage_context)
      except FileNotFoundError:
            # Index not found, create it
            reader = SimpleDirectoryReader(input_dir="./msft_cyber_docs")
            docs = reader.load_data()
            index = VectorStoreIndex.from_documents(docs, service_context=service_context)

            # Save the index to the file
            index.storage_context.persist(persist_dir=persist_dir)

      return index

index = load_data()

# DEFINE THE RUN_CHATS FUNCTION
def run_chats(query):
    
    search_time = 0.0 

    similarity_top_k = 16
   
    start_time_search = time.time() # time vector search
    chat_engine = index.as_chat_engine(chat_mode="condense_question",
                                       similarity_top_k=similarity_top_k, 
                                   ) # verbose=True # streaming=True
    end_time_search = time.time()
           
    result = chat_engine.chat(query)  # chat_engine.stream_chat(query)
    # Calculate search time
    search_time = end_time_search - start_time_search

    # Store the values of k
    query_params = similarity_top_k
    # result.print_response_stream()
    
    return result, query_params, search_time


# XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
# Create the selected version of the chat engine

if "chat_engine" not in st.session_state.keys(): # Initialize the chat engine
        st.session_state.chat_engine = None

if prompt := st.chat_input("Your question"): # Prompt for user input and save to chat history
    st.session_state.messages.append({"role": "user", "content": prompt})

for message in st.session_state.messages: # Display the prior chat messages
    with st.chat_message(message["role"]):
        st.write(message["content"])

# If last message is not from assistant, generate a new response
if st.session_state.messages[-1]["role"] != "assistant":
    with st.chat_message("assistant"):
        with st.spinner("Thinking..."):
            start_time = time.time()
            response, query_params, search_time = run_chats(st.session_state.messages[-1]["content"])
            response_time = time.time() - start_time
            st.write(response.response)
            message = {"role": "assistant", "content": response.response}
            st.session_state.messages.append(message) # Add response to message history
            # words = count_words(response.response)



