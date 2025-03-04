import requests
import time


class AIFeedbackProvider:
    def __init__(self, api_key):
        self.api_key = api_key

    def get_ai_feedback(self, prompt):
        try:
            headers = {
                "Authorization": "Bearer E7NOK7BCmDIHtNerhsWSJnyBJyL9YzAC",
                "Content-Type": "application/json",
            }

            payload = {
                "model": "open-mistral-7b",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 500,
                "temperature": 0.3,
            }

            response = requests.post(
                "https://api.mistral.ai/v1/chat/completions",
                headers=headers,
                json=payload,
            )
            response_json = response.json()

            # Handle API errors
            if "error" in response_json:
                return f"API Error: {response_json['error']}"

            # Check for rate limits
            if (
                "message" in response_json
                and "Requests rate limit exceeded" in response_json["message"]
            ):
                print("Rate limit hit, sleeping for 5 seconds...")
                time.sleep(5)  # Pause before retrying
                return self._get_ai_feedback(prompt)  # Retry the request

            # Extract feedback
            if "choices" in response_json and response_json["choices"]:
                return response_json["choices"][0]["message"]["content"].strip()
            else:
                return f"Error: Unexpected API response format: {response_json}"

        except Exception as e:
            return f"Error getting AI feedback: {str(e)}"
