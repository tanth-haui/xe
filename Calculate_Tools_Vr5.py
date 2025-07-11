import os
import ansa
from ansa import *

def ShowMessage(message, msg_type=guitk.constants.BCMessageBoxInformation, accept_text="OK", reject_text=None):
    msg_win = guitk.BCMessageWindowCreate(msg_type, message, True)
    guitk.BCMessageWindowSetAcceptButtonText(msg_win, accept_text)
    if reject_text:
        guitk.BCMessageWindowSetRejectButtonText(msg_win, reject_text)
    guitk.BCMessageWindowSetTextAlignment(msg_win, guitk.constants.BCAlignTop | guitk.constants.BCAlignHCenter)
    guitk.BCMessageWindowExecute(msg_win)

def SheetExists(ex_ref, sheet_name):
    try:
        _ = utils.XlsxGetCellValue(ex_ref, sheet_name, 1, 1)
        return True
    except:
        return False

def CalculateMassTool():
    TopWindow = guitk.BCWindowCreate("Calculate Mass Tools", guitk.constants.BCOnExitDestroy)
    guitk.BCWindowSetInitGeometry(TopWindow, 100, 100, 300, 200)

    BCButtonGroup_1 = guitk.BCButtonGroupCreate(TopWindow, "Input data info:", guitk.constants.BCVertical)

    guitk.BCLabelCreate(BCButtonGroup_1, "Base Model: ")
    base_file = guitk.BCLineEditPathCreate(BCButtonGroup_1, guitk.constants.BCHistoryFiles, "", guitk.constants.BCHistorySelect, "")

    guitk.BCLabelCreate(BCButtonGroup_1, "Excel File: ")
    excel_file = guitk.BCLineEditPathCreate(BCButtonGroup_1, guitk.constants.BCHistoryFiles, "", guitk.constants.BCHistorySelect, "")

    guitk.BCLabelCreate(BCButtonGroup_1, "Tên Xe:")
    name_car = guitk.BCLineEditCreate(BCButtonGroup_1, "")
    guitk.BCLineEditSetPlaceholderText(name_car, "Nhập tên xe cần đo")

    guitk.BCDialogButtonBoxCreate(TopWindow)
    data = (TopWindow, base_file, excel_file, name_car)

    guitk.BCWindowSetRejectFunction(TopWindow, CancelClickFunc, data)
    guitk.BCWindowSetAcceptFunction(TopWindow, OkClickFunc, data)
    guitk.BCShow(TopWindow)

def CancelClickFunc(TopWindow, data):
    return 1

def Calculate(pid_list):
    entities = []
    for pid in pid_list:
        entity = base.GetEntity(constants.NASTRAN, "__PROPERTIES__", pid)
        if entity:
            entities.append(entity)
    if entities:
        return base.DeckMassInfo(apply_on="custom", custom_entities=entities)
    return None

def ReadExcelInput(ExPath, sheet_name):
    ExRef = utils.XlsxOpen(ExPath)
    cot_dict = {}
    cot_map = {3: "D", 4: "E", 5: "F", 6: "G", 7: "H", 8: "I"}

    for col in range(3, 9):
        temp_col = []
        for row in range(10, 700):
            val = utils.XlsxGetCellValue(ExRef, sheet_name, row, col)
            if val == "" or val is None:
                break
            temp_col.append(val)

        col_values = [int(val.strip()) for val in temp_col if str(val).strip().isdigit()]
        col_name = f"COT_{cot_map[col]}"
        cot_dict[col_name] = col_values

    utils.XlsxClose(ExRef)
    return (cot_dict.get("COT_D", []), cot_dict.get("COT_E", []),
            cot_dict.get("COT_F", []), cot_dict.get("COT_G", []),
            cot_dict.get("COT_H", []), cot_dict.get("COT_I", []))

def OkClickFunc(TopWindow, data):
    link_base = guitk.BCLineEditPathSelectedFilePaths(data[1])
    link_excel = guitk.BCLineEditPathSelectedFilePaths(data[2])
    name_xe = guitk.BCLineEditGetText(data[3])

    if not link_base or not link_excel or not name_xe:
        ShowMessage("Hãy chọn đầy đủ Base, Excel và nhập tên xe", guitk.constants.BCMessageBoxWarning)
        return 0

    if not os.path.exists(link_base):
        ShowMessage(f"Không tìm thấy file Base tại: {link_base}", guitk.constants.BCMessageBoxCritical)
        return 0

    if not os.path.exists(link_excel) or not link_excel.endswith(".xlsx"):
        ShowMessage(f"File Excel không tồn tại hoặc không đúng định dạng: {link_excel}", guitk.constants.BCMessageBoxCritical)
        return 0

    ExRef = utils.XlsxOpen(link_excel)
    if not SheetExists(ExRef, name_xe):
        ShowMessage(f"Không tìm thấy sheet có tên '{name_xe}' trong file Excel.", guitk.constants.BCMessageBoxCritical)
        utils.XlsxClose(ExRef)
        return 0
    utils.XlsxClose(ExRef)

    pid_lists = ReadExcelInput(link_excel, name_xe)
    pid_lists_dict = {
        "List_BIW": pid_lists[0],
        "List_ENKON": pid_lists[1],
        "List_FR FLOOR": pid_lists[2],
        "List_RR FLOOR": pid_lists[3],
        "List_BODY": pid_lists[4],
        "List_torimu": pid_lists[5]
    }

    list_to_col_map = {
        "List_BIW": 3,
        "List_ENKON": 4,
        "List_FR FLOOR": 5,
        "List_RR FLOOR": 6,
        "List_BODY": 7,
        "List_torimu": 8
    }

    base.InputNastran(link_base, model_action="overwrite_model", properties_id="keep-new")

    ExRef = utils.XlsxOpen(link_excel)
    error_messages = []
    success = True

    for list_name, pid_list in pid_lists_dict.items():
        if not pid_list:
            error_messages.append(f"- {list_name}: Danh sách PID rỗng.")
            success = False
            continue

        total_mass = Calculate(pid_list)
        if total_mass:
            mass_values = float(f"{total_mass.net_mass:.4f}")
            print(f"{list_name} mass: {mass_values}")
            col = list_to_col_map[list_name]
            utils.XlsxSetCellValue(ExRef, name_xe, 6, col, str(mass_values))
        else:
            error_messages.append(f"- {list_name}: Không thể tính khối lượng.")
            success = False

    if success:
        utils.XlsxSave(ExRef, link_excel)
        ShowMessage("Đã tính xong khối lượng và lưu vào Excel.", guitk.constants.BCMessageBoxInformation)
    else:
        ShowMessage("Một số lỗi đã xảy ra, file Excel sẽ không được lưu:\n\n" + "\n".join(error_messages), guitk.constants.BCMessageBoxWarning)

    utils.XlsxClose(ExRef)

CalculateMassTool()
