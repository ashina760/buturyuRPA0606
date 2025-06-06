from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
from settings import CHROME_PATH, CHROMEDRIVER_PATH, AEON_OPCD, AEON_PASSWORD
from selenium.common.exceptions import TimeoutException

class AeonUploader:
    def __init__(self):
        self.chrome_path = CHROME_PATH
        self.driver_path = CHROMEDRIVER_PATH
        self.opcd = AEON_OPCD
        self.password = AEON_PASSWORD
        self.driver = None
        
    def setup_browser(self):
        options = Options()
        options.binary_location = self.chrome_path
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--start-maximized")
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')

        service = Service(self.driver_path)
        self.driver = webdriver.Chrome(service=service, options=options)

    def login(self):
        self.driver.get("1")
        self.driver.find_element(By.NAME, "OPCD").send_keys(self.opcd, Keys.RETURN)
        self.driver.find_element(By.NAME, "PSWD").send_keys(self.password)
        self.driver.find_element(By.XPATH, "//button[text()='ログイン']").click()

    def navigate_to_upload_page(self):
        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".tail_item_row_1:nth-child(6)"))
        ).click()

        self.driver.switch_to.window(self.driver.window_handles[1])

        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "button-1041-btnIconEl"))
        ).click()

        action = ActionChains(self.driver)
        element = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//span[text()='発注']"))
        )
        action.move_to_element(element).perform()

        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "menuitem-1049"))
        ).click()
        sleep(5)
    def upload_file(self, file_path):

        self.driver.find_element(By.ID, "filefield-1495-button-fileInputEl").send_keys(file_path)
        self.driver.find_element(By.ID, "ext-comp-1483cmdExec-btnIconEl").click()

        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "button-1006-btnIconEl"))
        ).click()
        sleep(2)

        try:
            WebDriverWait(self.driver, 5).until(
                EC.visibility_of_element_located((By.ID, "component-1002"))
            )
            # print(123123123)
            return {"error": "inputEl"} # 找到了，返回错误信息
        except TimeoutException:

            return {"error": False} 
        
    def extract_results(self):
        wait = WebDriverWait(self.driver, 15)

        print("🔍 正在查找表格行勾选框...")
        checkers = wait.until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "x-grid-row-checker"))
        )

        print(f"☑️ 共发现 {len(checkers)} 项可勾选")

        for i, checker in enumerate(checkers):
            try:
                # 滚动 + 使用 JS 方式点击（更强）
                self.driver.execute_script("arguments[0].scrollIntoView(true);", checker)
                self.driver.execute_script("arguments[0].click();", checker)
                # print(f"  ✅ 第 {i+1} 项已勾选")
                sleep(0.2)
            except Exception as e:
                print(f"  ❌ 第 {i+1} 项点击失败: {e}")

        # 点击下方按钮（用 JS 更稳）
        try:
            # print("📤 准备点击执行按钮...")
            # 找到 label 上写着 CSV 的元素
            label = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//label[text()='CSV']")
            ))

            # 通过 label 的 for 属性，反向获取对应的 input ID
            radio_id = label.get_attribute("for")
            radio_input = self.driver.find_element(By.ID, radio_id)

            # 使用 JavaScript 模拟点击，绕过 Ext JS 的阻挡
            self.driver.execute_script("arguments[0].scrollIntoView(true);", radio_input)
            self.driver.execute_script("arguments[0].click();", radio_input)

            # print("✅ 已点击 CSV radio")

            output_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='出力指示']"))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", output_button)
            self.driver.execute_script("arguments[0].click();", output_button)
            # print("✅ 已点击“出力指示”按钮")
            sleep(10)
        except TimeoutException:
            print("❌ 未找到执行按钮，点击失败")


    def run(self, file_path):
        try:
            self.setup_browser()
            self.login()
            # sleep(1000)
            self.navigate_to_upload_page()

            result = self.upload_file(file_path)
            if result.get("error") == "inputEl":
                sleep(5)
            
                return {"success": False, "inputEl": True,"error": "err_list"}
                
            self.extract_results()
            return {"success": True, "result": "pass"}

        except Exception as e:
            return {"success": False, "error": str(e)}
        
        finally:

            self.close()
    
    def close(self):
        if self.driver:
            self.driver.quit()
            print("✅ 浏览器已关闭")
        else:
            print("⚠️ 浏览器未正常启动，无需关闭")

