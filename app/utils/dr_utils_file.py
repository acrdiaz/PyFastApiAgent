import os


class DRUtilsFile:
    def __init__(self, file_path):
        self.file_path = file_path

    def get_file_size(self):
        try:
            return os.path.getsize(self.file_path)
        except FileNotFoundError:
            print(f"{self.file_path} file not found.")
            return 0

    def load_file(self):
        try:
            with open(self.file_path, 'r') as file:
                content = file.read().strip()
                if content:
                    return content
                else:
                    return None
        except FileNotFoundError:
            print(f"{self.file_path} file not found.")
            return None

    def write_file(self, prompt: str):
        try:
            with open(self.file_path, "w") as file:
                file.write(prompt)
        except FileNotFoundError:
            print(f"{self.file_path} file not found.")
        except Exception as e:
            print(f"An error occurred while writting to the prompt file: {e}")

    def clean_file(self):
        try:
            with open(self.file_path, "w") as file:
                file.write("")
        except FileNotFoundError:
            print(f"{self.file_path} file not found.")
        except Exception as e:
            print(f"An error occurred during cleaning the prompt file: {e}")