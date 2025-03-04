class PitchDeckReviewer:
    def __init__(self, feedback_provider):
        self.slides_content = []
        self.feedback = {}
        self.feedback_provider = feedback_provider

    def analyze_pitch_deck(self):
        """Analyzes the extracted text using an online AI API."""
        try:
            if not self.slides_content:
                return {"error": "No content extracted from slides."}

            structure_prompt = self._create_structure_prompt()
            self.feedback["structure"] = self.feedback_provider.get_ai_feedback(
                structure_prompt
            )
            self.feedback["slides"] = []

            for slide in self.slides_content:
                slide_prompt = self._create_slide_prompt(slide)
                slide_feedback = self.feedback_provider.get_ai_feedback(slide_prompt)
                self.feedback["slides"].append(
                    {"slide_number": slide["slide_number"], "feedback": slide_feedback}
                )

            recommendations_prompt = self._create_recommendations_prompt()
            self.feedback["recommendations"] = self.feedback_provider.get_ai_feedback(
                recommendations_prompt
            )

            return self.feedback
        except Exception as e:
            print(f"Error analyzing pitch deck: {e}")
            return {"error": f"Error analyzing pitch deck: {str(e)}"}

    def _create_structure_prompt(self):
        """Generates a prompt to analyze the overall pitch deck structure."""
        slides_overview = "\n\n".join(
            [
                f"Slide {slide['slide_number']}: {slide['text'][:200]}..."
                for slide in self.slides_content
            ]
        )
        return f"Analyze the structure of the pitch deck:\n{slides_overview}"

    def _create_slide_prompt(self, slide):
        """Generates a prompt for individual slide analysis."""
        return f"Analyze Slide {slide['slide_number']}:\n{slide['text']}"

    def _create_recommendations_prompt(self):
        """Generates a prompt for improvement recommendations."""
        structure_feedback = self.feedback.get("structure", "")
        slide_feedbacks = "\n\n".join(
            [
                f"Slide {item['slide_number']} Feedback: {item['feedback']}"
                for item in self.feedback.get("slides", [])
            ]
        )
        return f"Based on this analysis, provide key improvements:\n{structure_feedback}\n{slide_feedbacks}"
