from document_processor import DocumentProcessor, Config, WeChatDrive
import os
import shutil


class MockWeChatDrive(WeChatDrive):
    def start(self):
        print("Mock Drive: Started")

    def login(self):
        print("Mock Drive: Login simulated")

    def download_image(self, url, save_path):
        print(f"Mock Drive: Downloading {url} to {save_path}")
        # Create a dummy image for testing
        from PIL import Image
        img = Image.new('RGB', (100, 30), color=(73, 109, 137))
        img.save(save_path)
        return True

    def close(self):
        print("Mock Drive: Closed")


class TestProcessor(DocumentProcessor):
    def __init__(self):
        super().__init__()
        self.drive = MockWeChatDrive()


if __name__ == "__main__":
    # Ensure output directory is clean for testing
    if os.path.exists(Config.OUTPUT_BASE_DIR):
        shutil.rmtree(Config.OUTPUT_BASE_DIR)

    processor = TestProcessor()
    processor.run()
