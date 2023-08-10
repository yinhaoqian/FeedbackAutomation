from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

sourceFile = r"D:\Yinhao\userIds.txt"
destFile = r"D:\Yinhao\modelNumbers.txt"


def queryAll(UserIds):
    driver = webdriver.Chrome()
    Models = []
    driver.get("http://sgp.admin.iot.mi.srv/logQuery/userInformation")
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ant-input"))
    )
    print("login detected lah :)")
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder,'uid')]"))
    )
    for userId in UserIds:
        print("***Querying userId = " + str(userId))
        uidInput = driver.find_element(By.XPATH, "//input[contains(@placeholder,'uid')]")
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder,'uid')]")))
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//button[@_nk='S5Uo11']")))
        # time.sleep(1)
        uidInput.send_keys(Keys.CONTROL + "a")
        uidInput.send_keys(Keys.DELETE)
        # time.sleep(1)
        uidInput.send_keys(userId)
        uidInput.send_keys(Keys.ENTER)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//button[@_nk='S5Uo51']")))
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//button[@_nk='S5Uo11']")))
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//div[@_nk='IbVG11']")))
            ModelNumbers = driver.find_elements(By.XPATH, "//div[@_nk='IbVG11']")
            Result = []
            print(len(ModelNumbers))
            for modelNumber in ModelNumbers:
                print(modelNumber.text)
                Result.append(modelNumber.text)
            Models.append("|".join(Result))
        except:
            print(" - ")
            Models.append(" - ")
    return Models


def processExcel(excelFilePath):
    excel = openpyxl.load_workbook(excelFilePath, data_only=True)
    sheet = excel["FEEDBACK_REF"]
    userIds = []
    for x in range(2, 1024):
        if not sheet["BD" + str(x)].value:
            break
        userIds.append(sheet["BD" + str(x)].value)
    Models = queryAll(userIds)
    for i in range(0, len(userIds)):
        feedbackCode = sheet["A" + str(i + 2)].value
        deviceId = ""
        Models_ = Models[i].split("|")
        if len(Models_) > 1:
            for model in Models_:
                if 20 <= feedbackCode <= 34 and "camera" in model:
                    deviceId = model
                    break
                elif 35 <= feedbackCode <= 59 and "vacuum" in model:
                    deviceId = model
                    break
                elif 60 <= feedbackCode <= 69 and "scooter" in model:
                    deviceId = model
                    break
                elif 70 <= feedbackCode <= 80 and "repeater" in model:
                    deviceId = model
                    break
        if deviceId == "":
            deviceId = Models_[0]
        deviceType = ""
        if "camera" in deviceId:
            deviceType = "摄像机"
        elif "vacuum" in deviceId:
            deviceType = "扫地机"
        elif "scooter" in deviceId:
            deviceType = "滑板车"
        elif "repeater" in deviceId:
            deviceType = "WiFi信号放大器"
        else:
            deviceType = "其他"
        #sheet["E" + str(i + 2)].value = deviceType
        sheet["F" + str(i + 2)].value = deviceId
        print("processExcel -> "+str(feedbackCode)+" "+str(deviceType)+" "+str(deviceId))
    excel.save(excelFilePath+".xlsx")
    i = 2


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    processExcel("C:/Users/MI/Documents/TEMP.xlsx")
