import pandas as pd
from rapidfuzz import process, fuzz


class AdmissionChatbot:
    def __init__(self, excel_path):
        self.knowledge_base = []
        sheets = ["Admission and Registration", "Admission & Registration 2", "Programs", "Orientation"]

        df_dict = pd.read_excel(excel_path, sheet_name=sheets)

        for sheet, df in df_dict.items():
            # Normalize column names
            df.columns = df.columns.str.strip().str.lower()
            question_col = None
            answer_col = None

            for col in df.columns:
                if 'question' in col:
                    question_col = col
                elif 'answer' in col:
                    answer_col = col

            if question_col and answer_col:
                for _, row in df.iterrows():
                    question = str(row[question_col]).strip()
                    answer = str(row[answer_col]).strip()

                    if question and answer:
                        # Add original QA pair
                        self.knowledge_base.append((question.lower(), answer))

                        # Add variations for better matching
                        self._add_alternates(question, answer)

        print(f"Loaded {len(self.knowledge_base)} QA pairs from all sheets.")

    def _add_alternates(self, question, answer):
        question = question.lower()
        if "apply" in question or "application" in question:
            self.knowledge_base.append(("how do i apply for admission", answer))
            self.knowledge_base.append(("how to apply", answer))
            self.knowledge_base.append(("application process", answer))
        if "requirement" in question or "requirements" in question:
            self.knowledge_base.append(("what are the admission requirements", answer))
            self.knowledge_base.append(("admission criteria", answer))
        if "financial aid" in question or "scholarship" in question:
            self.knowledge_base.append(("is financial aid available", answer))
            self.knowledge_base.append(("does kepler offer financial aid", answer))
            self.knowledge_base.append(("are there any scholarships", answer))
            self.knowledge_base.append(("is there any scholarship", answer))
        if "deadline" in question:
            self.knowledge_base.append(("what is the application deadline", answer))
        if "transfer" in question or "credits" in question:
            self.knowledge_base.append(("can i transfer credits from another institution", answer))
        if "contact" in question or "email" in question or "phone" in question:
            self.knowledge_base.append(("how can i contact the admissions office", answer))
        if "program" in question or "degree" in question:
            self.knowledge_base.append(("what programs does kepler college offer", answer))
            self.knowledge_base.append(("what kind of degrees does kepler offer", answer))
        if "internship" in question or "practical training" in question:
            self.knowledge_base.append(("does kepler college offer internships or practical training", answer))
        if "housing" in question or "dormitory" in question:
            self.knowledge_base.append(("does kepler college offer housing to students", answer))
        if "career" in question or "job" in question:
            self.knowledge_base.append(("what are kepler college graduates job prospects", answer))
            self.knowledge_base.append(("career outcomes", answer))
        if "teaching" in question or "approach" in question:
            self.knowledge_base.append(("what is kepler college's teaching approach", answer))
        if "technology" in question or "learning environment" in question:
            self.knowledge_base.append(("how does kepler college use technology in its learning environment", answer))
        if "international" in question:
            self.knowledge_base.append(("is the education at kepler college recognized internationally", answer))
        if "language" in question or "instruction" in question:
            self.knowledge_base.append(("what languages are used for instruction at kepler college", answer))
        if "alumni" in question or "networking" in question:
            self.knowledge_base.append(("what opportunities are there for alumni engagement at kepler college", answer))
        if "study abroad" in question or "exchange" in question:
            self.knowledge_base.append(("are there any study abroad programs or exchange opportunities", answer))
        if "research" in question:
            self.knowledge_base.append(("what kind of research opportunities are available at kepler", answer))
        if "core value" in question or "mission" in question:
            self.knowledge_base.append(("what is kepler college's mission", answer))
            self.knowledge_base.append(("what are kepler’s core values", answer))
        if "intake" in question:
            self.knowledge_base.append(("how many intakes", answer))
            self.knowledge_base.append(("when is intake open", answer))
            self.knowledge_base.append(("intake open", answer))
        if "support" in question or "services" in question:
            self.knowledge_base.append(("what support services are available for students", answer))
        if "campus" in question:
            self.knowledge_base.append(("what is the campus like", answer))
            self.knowledge_base.append(("what are the campuses", answer))
            self.knowledge_base.append(("campuses", answer))
        if "student life" in question or "life on campus" in question:
            self.knowledge_base.append(("what is life like on kepler college's campus", answer))
        if "flexible payment" in question or "payment plan" in question:
            self.knowledge_base.append(("are there flexible payment plans for tuition", answer))
        if "leadership" in question or "extracurricular" in question:
            self.knowledge_base.append(("what leadership or extracurricular opportunities does kepler college offer", answer))
        if "volunteer" in question or "community" in question:
            self.knowledge_base.append(("can i engage in volunteer work or community projects while at kepler college", answer))
        if "diverse" in question or "inclusion" in question:
            self.knowledge_base.append(("how diverse is kepler college college's student body", answer))
            self.knowledge_base.append(("how does kepler college handle student diversity and inclusion", answer))
        if "difference" in question or "unique" in question:
            self.knowledge_base.append(("what makes kepler college different from other universities in rwanda", answer))
        if "entrepreneurship" in question:
            self.knowledge_base.append(("how does kepler college support student entrepreneurship", answer))
        if "online learning" in question or "distance learning" in question:
            self.knowledge_base.append(("does kepler college have an online learning option", answer))
        if "postgraduate" in question or "master" in question:
            self.knowledge_base.append(("does kepler offer postgraduate programs", answer))
        if "technology resource" in question or "software" in question:
            self.knowledge_base.append(("what are the technological resources available to students", answer))
        if "balance" in question or "personal life" in question:
            self.knowledge_base.append(("how does kepler college help students balance academic and personal life", answer))
        if "leadership development" in question:
            self.knowledge_base.append(("what is kepler college's approach to leadership development", answer))
        if "social impact" in question:
            self.knowledge_base.append(("how does kepler college integrate social impact into its education", answer))
        if "sustainability" in question:
            self.knowledge_base.append(("what is kepler's approach to sustainability and community engagement", answer))
        if "faculty selection" in question:
            self.knowledge_base.append(("how are faculty members selected at kepler college", answer))
        if "partnership" in question or "university" in question:
            self.knowledge_base.append(("are there any partnerships between kepler and other universities", answer))
        if "kepler college" in question or "about kepler" in question:
            self.knowledge_base.append(("tell me about kepler college", answer))
        if "student to faculty ratio" in question or "ratio" in question:
            self.knowledge_base.append(("what is the student-to-faculty ratio at kepler", answer))
        if "career center" in question or "career services" in question:
            self.knowledge_base.append(("is there a career center at kepler", answer))
        if "academic support" in question or "struggling" in question:
            self.knowledge_base.append(("what kind of academic support is available for students struggling with their studies", answer))
        if "innovation" in question:
            self.knowledge_base.append(("how does kepler foster innovation among its students", answer))
        if "research" in question:
            self.knowledge_base.append(("what kind of research opportunities are available at kepler", answer))
        if "challenges" in question:
            self.knowledge_base.append(("what are the key challenges kepler students face, and how does the college address them", answer))
        if "education landscape" in question:
            self.knowledge_base.append(("how does kepler contribute to the education landscape in rwanda", answer))
        if "internships" in question:
            self.knowledge_base.append(("what kind of internships are available to kepler students", answer))
        if "relevant" in question or "job market" in question:
            self.knowledge_base.append(("how does kepler ensure its programs stay relevant to the job market", answer))
        if "alumni connection" in question or "network" in question:
            self.knowledge_base.append(("how do kepler students stay connected with alumni", answer))
        if "global job market" in question:
            self.knowledge_base.append(("how does kepler prepare students for the global job market", answer))

    def get_response(self, user_input):
        if not user_input.strip():
            return "Please ask a question."

        user_input = user_input.lower().strip()

        matches = process.extract(user_input, [q[0] for q in self.knowledge_base], scorer=fuzz.token_sort_ratio, limit=3)
        best_match = matches[0]

        if best_match[1] > 65:
            return self.knowledge_base[best_match[2]][1]
        else:
            return "I'm sorry, I couldn't find an answer to that. Please check Kepler College's website or contact admissions@keplercollege.ac.rw."