import streamlit as st
import pandas as pd
from io import BytesIO

st.title("PragyanAI FAQ Excel Generator")


faq_data = {
    "Category": [
        "Program Overview", "Program Structure", "Program Structure",
        "Pricing & Fees", "Pricing & Fees", "Curriculum & Skills",
        "Curriculum & Skills", "Evaluation & Projects", "Career & Placement",
        "Leadership & Contact"
    ],
    "Question": [
        "What is the total duration and structure of the PragyanAI program?",
        "What happens in Phase 1 (First 6 Months)?",
        "What happens in Phase 2 (12 Months)?",
        "What is the fee structure for the Founding Batch?",
        "What is the salary potential after completing the program?",
        "What modules are covered in Months 1-3 (Foundational Core)?",
        "What modules are covered in Months 4-6 (Advanced Frontier)?",
        "How are students evaluated during the 6-month offline training?",
        "What career tracks or roles are unlocked?",
        "Who leads PragyanAI and how can I contact them?"
    ],
    "Answer": [
        "The PragyanAI AI GenAI program is an 18-month journey comprising 6 Months of Fully Offline Training followed by a 12-Month Internship & Placement Drive.",
        "Phase 1 (6 Months) consists of intensive offline training with half-day classroom sessions, half-day hands-on labs, real-time projects, monthly hackathons, and technical seminars.",
        "Phase 2 (12 Months) includes an extended internship, live client assignments, technical mock interviews, resume building, and startup/product development exposure.",
        "Founding Batch (First 100 students): Initial Training Fee is ₹50,000 + Success Fee of ₹50,000 after placement (Total ₹1,00,000, discounted from standard ₹1,50,000).",
        "Target packages: AI Engineer (₹6–₹15 LPA), GenAI Engineer (₹8–₹18 LPA), and Agentic AI Engineer (₹10–₹25 LPA).",
        "Month 1: Python Full Stack & Analytics. Month 2: Data Science & BI Analytics. Month 3: Machine Learning Frameworks (AutoML, Streamlit deployment).",
        "Month 4: Deep Learning & Computer Vision (CNNs, PyTorch, YOLO). Month 5: NLP & Generative AI (LLMs, RAG, LangChain, Fine-tuning). Month 6: Agentic AI (CrewAI, AutoGen, Multi-Agent Systems, MCP).",
        "Students participate in 1 Technical Seminar per skill (evaluated out of 100 marks) and 1 Skill-wise 48-Hour Hackathon with cash prizes (₹5,000 winner, ₹3,000 runner-up).",
        "7 Multi-Track Pathways: Data Analyst, Data Scientist & ML, AI Engineer, GenAI Engineer, Agentic AI Engineer, Product/MVP Engineer, and Software Engineer.",
        "Led by Sateesh Ambesange (Co-Founder, NITK alumnus, 25+ years IT exp). Phone: +91-9741007422 | Email: sateesh.ambesange@pragyanai.com / pragyan.ai.school@gmail.com"
    ]
}

df = pd.DataFrame(faq_data)

# Display the DataFrame
st.subheader("PragyanAI FAQ Data")
st.dataframe(df, use_container_width=True)

# Create Excel file in memory
output = BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="PragyanAI FAQ")

output.seek(0)

# Download button
st.download_button(
    label="📥 Download PragyanAI FAQ Excel",
    data=output,
    file_name="pragyan_faq_prices.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
import os
import streamlit as st
import pandas as pd

# LangChain
from langchain_community.document_loaders import PyPDFLoader
from langchain_community.vectorstores import FAISS
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_core.documents import Document
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.output_parsers import StrOutputParser
from langchain_community.chat_message_histories import ChatMessageHistory
from langchain_core.runnables.history import RunnableWithMessageHistory

# Groq
from langchain_groq import ChatGroq
# ---------------------------------------------------------------------------
# 1. System Prompts specifically grounded in PragyanAI Data
# ---------------------------------------------------------------------------
SALES_PROMPTS = {
    "PragyanAI Student Counselor": """You are Aarav, an Academic & Career Advisor for PragyanAI.
Goal: Guide prospective students to enroll in the 18-Month AI/GenAI Program (6 Month Offline Training + 12 Month Placement Drive).

Strict Rule: Answer pricing, fee structures, curriculum details, and salary potential ONLY based on the Document Context below.

Retrieved Document Context:
{context}

Behavior Guidelines:
1. Be encouraging, empathetic, and focus on practical "builder" skill transformation.
2. Highlight key advantages: 100+ projects, 48-hour hackathons, risk-shared pricing (pay-after-placement success fee), and direct mentorship under Sateesh Ambesange.""",

    "PragyanAI Student Counselor": """You are Aarav, an Academic & Career Advisor for PragyanAI.
Goal: Guide prospective students to enroll in the 18-Month AI/GenAI Program (6 Month Offline Training + 12 Month Placement Drive).

Strict Rule: Answer pricing, fee structures, curriculum details, and salary potential ONLY based on the Document Context below.

Retrieved Document Context:
{context}

Behavior Guidelines:
1. Be encouraging, empathetic, and focus on practical "builder" skill transformation.
2. Highlight key advantages: 100+ projects, 48-hour hackathons, risk-shared pricing (pay-after-placement success fee), and direct mentorship under Sateesh Ambesange.""",

    "PragyanAI Institutional / CoE Advisor": """You are Dr. Kavita, Institutional Relations Lead at PragyanAI.
Goal: Partner with engineering colleges to solve the education trap and transform students from theory learners into product builders.

Strict Rule: Use the retrieved Context below to cite exact program structures, multi-track career pathways, and evaluation rubrics (seminars, hackathons).

Retrieved Document Context:
{context}

Behavior Guidelines:
1. Maintain an authoritative, industry-oriented tone.
2. Focus on bridging the gap between college curricula and high-value industry roles (Agentic AI, GenAI).""",

    "PragyanAI Enterprise AI & Placement Lead": """You are Rohan, Enterprise Placement & Venture Lead at PragyanAI.
Goal: Connect hiring partners and enterprise leaders with top-tier PragyanAI builders and discuss talent recruitment or custom AI automation.

Strict Rule: Reference exact technical skills (CrewAI, AutoGen, LangChain, RAG, Multi-Agent systems) and portfolio deliverables (GitHub profile, live deployed MVPs) from the context below.

Retrieved Document Context:
{context}

Behavior Guidelines:
1. Confident, direct, and ROI-driven tone.
2. Emphasize that PragyanAI engineers are class-hired builders capable of deploying live applications immediately."""
}
import os
import tempfile
import pandas as pd
import streamlit as st

from langchain_community.document_loaders import PyPDFLoader
from langchain_community.vectorstores import FAISS
from langchain_huggingface import HuggingFaceEmbeddings
from langchain_core.documents import Document

# ---------------------------------------------------------------------------
# 2. Vector Store Indexer (Loads Excel FAQ + PDF Documents)
# ---------------------------------------------------------------------------

embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
vectorstore = None


def load_documents_into_vectorstore(uploaded_files=None):
    global vectorstore

    docs = []

    # ---------------------------
    # Process uploaded files
    # ---------------------------
    if uploaded_files:
        for uploaded_file in uploaded_files:

            suffix = os.path.splitext(uploaded_file.name)[1]

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                tmp_file.write(uploaded_file.getbuffer())
                temp_path = tmp_file.name

            if suffix.lower() == ".pdf":
                loader = PyPDFLoader(temp_path)
                docs.extend(loader.load())

            elif suffix.lower() in [".xlsx", ".xls"]:
                excel_df = pd.read_excel(temp_path)

                for _, row in excel_df.iterrows():
                    content = " | ".join(
                        [f"{col}: {val}" for col, val in row.items()]
                    )

                    docs.append(
                        Document(
                            page_content=content,
                            metadata={"source": uploaded_file.name}
                        )
                    )

    # ---------------------------
    # Load default FAQ automatically
    # ---------------------------
    if os.path.exists("pragyan_faq_prices.xlsx"):

        excel_df = pd.read_excel("pragyan_faq_prices.xlsx")

        for _, row in excel_df.iterrows():

            content = " | ".join(
                [f"{col}: {val}" for col, val in row.items()]
            )

            docs.append(
                Document(
                    page_content=content,
                    metadata={"source": "pragyan_faq_prices.xlsx"}
                )
            )

    # ---------------------------
    # Fallback knowledge
    # ---------------------------
    if not docs:

        docs = [
            Document(
                page_content="PragyanAI Program: 6 Months Offline Training + 12 Months Placement Drive. Led by Sateesh Ambesange."
            ),
            Document(
                page_content="Founding Batch Fee: ₹50,000 initial training + ₹50,000 success fee after placement."
            )
        ]

    vectorstore = FAISS.from_documents(docs, embeddings)

    return f"✅ Knowledge Base updated with {len(docs)} document chunks!"
import streamlit as st
from langchain_groq import ChatGroq
from langchain_core.prompts import ChatPromptTemplate, MessagesPlaceholder
from langchain_core.output_parsers import StrOutputParser
from langchain_community.chat_message_histories import ChatMessageHistory

# ---------------------------------------------------------------------------
# 3. Groq LLM & LCEL RAG Pipeline
# ---------------------------------------------------------------------------

# API Key
GROQ_BEC_API_KEY = "your_groq_bec_api_key"
# or:
# groq_bec_api_key = os.getenv("GROQ_BEC_API_KEY")

# LLM
llm = ChatGroq(
    groq_api_key=groq_bec_api_key,
    model_name="llama-3.3-70b-versatile",
    temperature=0.3
)

# Session Store
if "store" not in st.session_state:
    st.session_state.store = {}

def get_session_history(session_id: str):
    if session_id not in st.session_state.store:
        st.session_state.store[session_id] = ChatMessageHistory()

    return st.session_state.store[session_id]


def create_rag_chain(persona_name: str, retrieved_context: str):

    system_instruction = SALES_PROMPTS.get(
        persona_name,
        SALES_PROMPTS["PragyanAI Student Counselor"]
    ).format(context=retrieved_context)

    prompt = ChatPromptTemplate.from_messages([
        ("system", system_instruction),
        MessagesPlaceholder(variable_name="history"),
        ("human", "{input}")
    ])

    return prompt | llm | StrOutputParser()
    st.title("🤖 PragyanAI RAG Chatbot")

persona_name = st.sidebar.selectbox(
    "Choose Assistant",
    list(SALES_PROMPTS.keys())
)

if "messages" not in st.session_state:
    st.session_state.messages = []

# Display previous messages
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# User input
if prompt := st.chat_input("Ask anything about PragyanAI..."):

    st.session_state.messages.append(
        {
            "role": "user",
            "content": prompt
        }
    )

    with st.chat_message("user"):
        st.markdown(prompt)

    answer = respond(prompt, persona_name)

    st.session_state.messages.append(
        {
            "role": "assistant",
            "content": answer
        }
    )

    with st.chat_message("assistant"):
        st.markdown(answer)

import streamlit as st

# ---------------------------------------------------------------------------
# Streamlit User Interface
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="PragyanAI Intelligent Assistant",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 PragyanAI Conversational Sales & FAQ Assistant")
st.write(
    "Answers program questions based on the **PragyanAI Presentation & FAQ Sheet**."
)

# ---------------- Sidebar ----------------
with st.sidebar:

    st.header("Settings")

    persona_selector = st.selectbox(
        "Select PragyanAI Persona",
        list(SALES_PROMPTS.keys()),
        index=0
    )

    uploaded_files = st.file_uploader(
        "Upload Additional PDFs or Excel Sheets",
        type=["pdf", "xlsx", "xls"],
        accept_multiple_files=True
    )

    if st.button("📚 Update Knowledge Base"):

        status = load_documents_into_vectorstore(uploaded_files)
        st.success(status)

    if st.button("🗑 Clear Memory"):

        clear_chat_history(persona_selector)
        st.success("Chat history cleared.")
        st.rerun()

# ---------------- Chat History ----------------
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display previous messages
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ---------------- Chat Input ----------------
user_input = st.chat_input("Ask anything about PragyanAI...")

if user_input:

    # Show user message
    st.session_state.messages.append(
        {
            "role": "user",
            "content": user_input
        }
    )

    with st.chat_message("user"):
        st.markdown(user_input)

    # Generate response
    with st.spinner("Thinking..."):

        answer = respond(user_input, persona_selector)

    # Show assistant response
    st.session_state.messages.append(
        {
            "role": "assistant",
            "content": answer
        }
    )

    with st.chat_message("assistant"):
        st.markdown(answer)
