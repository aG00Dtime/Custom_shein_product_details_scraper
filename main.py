import os
import re
import sys
import threading

import requests
from PyQt5 import QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication, QVBoxLayout, QWidget, QTextEdit, QLabel
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from PyQt5.QtCore import QTimer


class MainWidget(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Drag and Drop")
        self.setFixedSize(500, 300)
        self.setAcceptDrops(True)

        self.output_window = QTextEdit()
        self.output_window.setReadOnly(True)
        self.output_window.setFixedHeight(100)

        self.label = QLabel("Drag file here")
        self.label.setAlignment(QtCore.Qt.AlignCenter)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        layout.addWidget(self.label)

        layout.addWidget(self.output_window)
        central_widget.setLayout(layout)

        QTimer.singleShot(0, self.scroll_output)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def scroll_output(self):
        scrollbar = self.output_window.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file_path in files:
            self.output_window.append(file_path)

            self.output_window.append("***** Grabbing Product Details *****")

            # scrape using thread
            threading.Thread(target=self.get_item_urls, args=(file_path,)).start()

    def get_item_data(self, url):
        # sizes and urls
        sizes = []
        picture_urls = []
        price = None

        chrome_options = Options()
        chrome_options.add_argument("--headless")  # Run Chrome in headless mode
        chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration
        driver = webdriver.Chrome(options=chrome_options)

        # open the page
        driver.get(url)

        # wait for it to finish loading, sometimes the images don't load
        driver.implicitly_wait(1)

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # get the product name
        # item_name_tag = soup.find('h1', class_='product-intro__head-name fsp-element')
        # item_name = item_name_tag.text if item_name_tag else None
        item_name = soup.title.text.strip()[:50]
        item_name = re.sub(r'[<>:"/\\|?*]', '_', item_name)
        item_name = item_name.replace(' ', '_')

        self.output_window.append(f"Grabbing: {item_name}")
        self.scroll_output()

        # extract the picture urls to download, not 100%
        picture_url_tag = soup.find_all('div', class_='product-intro-zoom__item')
        picture_thumbnails = soup.find_all('div', class_='product-intro__thumbs-item')

        for element in picture_url_tag:
            img_tag = element.find('img', class_='lazyload crop-image-container__img')
            if img_tag:
                url = img_tag['data-src']
                # remove  //
                if url.startswith("//"):
                    url = "https:" + url
                picture_urls.append(url)

        for element in picture_thumbnails:
            img_tag = element.find('img', class_='fsp-element crop-image-container__img')
            if img_tag:
                url = img_tag['src']
                # remove  //
                if url.startswith("//"):
                    url = "https:" + url
                picture_urls.append(url)

        size_elements = soup.find_all('div', class_='product-intro__size-choose')
        for size_element in size_elements:
            size_items = size_element.find_all('div', class_='product-intro__size-radio')
            for size_item in size_items:
                size = size_item.get('data-attr_value_name')
                if size:
                    sizes.append(size)

        # price, need to cater for discounts too
        elements = soup.find_all('div', class_='from original')
        for element in elements:
            if "$" in element.text:
                price = element.text
                break

        # folder name
        folder_name = item_name.replace(' ', '_').replace(',', '_').replace('&', '_')
        os.makedirs(folder_name, exist_ok=True)

        # concat sizes
        sizes_str = "_".join(sizes)

        # save the images in the urls
        for i, picture_url in enumerate(picture_urls):
            image_size = int(requests.head(picture_url).headers['Content-Length'])
            if image_size < 5 * 1024:
                continue

            filename = f"{item_name}_{price}_{sizes_str}_{i}.jpg"
            filepath = os.path.join(folder_name, filename)

            with open(filepath, 'wb') as f:
                f.write(requests.get(picture_url).content)

        # self.output_window.append(f"Done: {item_name}")

        driver.quit()

    def get_item_urls(self, file_name: str):
        with open(file_name, "r") as f:
            for item_url in f:
                item_url = item_url.strip()
                if not item_url:
                    continue
                self.get_item_data(item_url)

        self.output_window.append(f"Completed : {file_name}")

        self.scroll_output()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = MainWidget()
    ui.show()
    sys.exit(app.exec_())
