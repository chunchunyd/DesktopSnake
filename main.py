# -*- coding: utf-8 -*-
import win32gui
import win32con
import win32api
import win32process
import commctrl
import ctypes
from ctypes import wintypes
import time
import random
import keyboard  # 需要先安装: pip install keyboard
from collections import deque  # 引入双端队列


# --- 颜色定义 ---
class Colors:
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"


# --- 全局变量 ---
# 定义桌面窗口的类名
PROGMAN = "Progman"
SHELLDLL_DEFVIEW = "SHELLDLL_DefView"
SYS_LIST_VIEW32 = "SysListView32"

# 游戏状态
game_running = True
initial_positions = {}  # 存储图标的初始位置

# --- Win32/ctypes 定义 ---
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

kernel32.VirtualAllocEx.restype = wintypes.LPVOID
kernel32.VirtualAllocEx.argtypes = (
    wintypes.HANDLE,
    wintypes.LPVOID,
    ctypes.c_size_t,
    wintypes.DWORD,
    wintypes.DWORD,
)
kernel32.VirtualFreeEx.argtypes = (
    wintypes.HANDLE,
    wintypes.LPVOID,
    ctypes.c_size_t,
    wintypes.DWORD,
)
kernel32.ReadProcessMemory.argtypes = (
    wintypes.HANDLE,
    wintypes.LPCVOID,
    wintypes.LPVOID,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
)
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
kernel32.GetConsoleWindow = ctypes.windll.kernel32.GetConsoleWindow
kernel32.GetConsoleWindow.restype = wintypes.HWND
kernel32.GetConsoleWindow.argtypes = []


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


# --- 核心函数 ---
def find_desktop_listview_handle():
    try:
        h_progman = win32gui.FindWindow(PROGMAN, None)
        h_shelldll = win32gui.FindWindowEx(h_progman, 0, SHELLDLL_DEFVIEW, None)
        h_listview = win32gui.FindWindowEx(h_shelldll, 0, SYS_LIST_VIEW32, None)
        if not h_listview:
            print(f"{Colors.FAIL}错误: 找不到'SysListView32'窗口。{Colors.ENDC}")
            return None
        print(f"{Colors.OKGREEN}成功找到桌面ListView句柄: {h_listview}{Colors.ENDC}")
        return h_listview
    except Exception as e:
        print(f"{Colors.FAIL}寻找句柄时发生异常: {e}{Colors.ENDC}")
        return None


def get_icon_count(h_listview):
    if not h_listview:
        return 0
    return win32gui.SendMessage(h_listview, commctrl.LVM_GETITEMCOUNT, 0, 0)


def get_icon_position(h_listview, index):
    if not h_listview:
        return None, None
    tid, pid = win32process.GetWindowThreadProcessId(h_listview)
    h_process = kernel32.OpenProcess(
        win32con.PROCESS_VM_OPERATION
        | win32con.PROCESS_VM_READ
        | win32con.PROCESS_VM_WRITE,
        False,
        pid,
    )
    if not h_process:
        return None, None
    p_remote_buffer = None
    try:
        p_remote_buffer = kernel32.VirtualAllocEx(
            h_process,
            0,
            ctypes.sizeof(POINT),
            win32con.MEM_COMMIT | win32con.MEM_RESERVE,
            win32con.PAGE_READWRITE,
        )
        if not p_remote_buffer:
            return None, None
        result = win32gui.SendMessage(
            h_listview, commctrl.LVM_GETITEMPOSITION, index, p_remote_buffer
        )
        if result == 0:
            return None, None
        local_point = POINT()
        bytes_read = ctypes.c_size_t(0)
        kernel32.ReadProcessMemory(
            h_process,
            p_remote_buffer,
            ctypes.byref(local_point),
            ctypes.sizeof(local_point),
            ctypes.byref(bytes_read),
        )
        return local_point.x, local_point.y
    finally:
        if p_remote_buffer:
            kernel32.VirtualFreeEx(h_process, p_remote_buffer, 0, win32con.MEM_RELEASE)
        if h_process:
            kernel32.CloseHandle(h_process)


def set_icon_position(h_listview, index, x, y):
    if not h_listview:
        return
    lparam = win32api.MAKELONG(int(x), int(y))
    win32gui.SendMessage(h_listview, commctrl.LVM_SETITEMPOSITION, index, lparam)


# --- 游戏逻辑函数 ---


def save_initial_positions(h_listview, icon_count):
    global initial_positions
    # print(f"{Colors.OKCYAN}正在保存图标初始位置...{Colors.ENDC}")
    initial_positions = {}
    for i in range(icon_count):
        pos = get_icon_position(h_listview, i)
        if pos and pos[0] is not None:
            initial_positions[i] = pos
    # print(f"{Colors.OKGREEN}初始位置保存完毕。{Colors.ENDC}")


def restore_initial_positions(h_listview):
    global initial_positions
    if not initial_positions:
        return
    # print(f"{Colors.OKCYAN}正在恢复图标初始位置...{Colors.ENDC}")
    for index, (x, y) in initial_positions.items():
        set_icon_position(h_listview, index, x, y)
    # print(f"{Colors.OKGREEN}位置恢复完毕。{Colors.ENDC}")


def stop_game():
    global game_running
    print(f"\n{Colors.WARNING}接收到停止信号，游戏即将退出...{Colors.ENDC}")
    game_running = False


def calculate_grid_parameters(h_listview):
    """根据图标在桌面上的实际位置计算网格参数。"""
    print(f"{Colors.OKCYAN}正在扫描桌面图标以计算网格参数...{Colors.ENDC}")
    all_icon_pos = []
    icon_count = get_icon_count(h_listview)
    for i in range(icon_count):
        pos = get_icon_position(h_listview, i)
        if pos:
            all_icon_pos.append(pos)

    if len(all_icon_pos) < 4:
        print(f"{Colors.FAIL}错误：桌面上可识别的图标少于4个。{Colors.ENDC}")
        return None

    sorted_by_topleft = sorted(all_icon_pos, key=lambda p: p[0] + p[1])
    pos0 = sorted_by_topleft[0]
    origin_x, origin_y = pos0

    left_column_icons = sorted(
        [p for p in all_icon_pos if abs(p[0] - origin_x) < 20], key=lambda p: p[1]
    )
    if len(left_column_icons) < 2:
        print(
            f"{Colors.FAIL}错误：无法找到左下角边界图标。请确保在左上角图标的正下方放置一个图标作为边界。{Colors.ENDC}"
        )
        return None
    pos3 = left_column_icons[-1]

    pos1 = None
    pos2 = None
    min_dx = float("inf")
    min_dy = float("inf")

    for x, y in all_icon_pos:
        if x == origin_x and y == origin_y:
            continue
        dx = x - origin_x
        dy = y - origin_y
        if abs(dy) < 20 and 0 < dx < min_dx:
            min_dx = dx
            pos1 = (x, y)
        if abs(dx) < 20 and 0 < dy < min_dy:
            min_dy = dy
            pos2 = (x, y)

    if not pos1 or not pos2:
        print(f"{Colors.FAIL}错误：无法自动确定网格大小。{Colors.ENDC}")
        print(
            f"{Colors.FAIL}请确保在左上角图标的 正右方 和 正下方 紧邻的位置有其他图标。{Colors.ENDC}"
        )
        return None

    grid_size_x = pos1[0] - pos0[0]
    grid_size_y = pos2[1] - pos0[1]

    if grid_size_x <= 0 or grid_size_y <= 0:
        print(f"{Colors.FAIL}错误：网格计算失败。请检查参考图标位置。{Colors.ENDC}")
        return None

    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    cols = (screen_width - origin_x) // grid_size_x
    rows = (pos3[1] - origin_y) // grid_size_y + 1

    grid_info = {
        "size_x": grid_size_x,
        "size_y": grid_size_y,
        "origin_x": origin_x,
        "origin_y": origin_y,
        "cols": cols,
        "rows": rows,
    }

    print(f"{Colors.OKGREEN}网格参数计算成功:{Colors.ENDC}")
    print(f"  - 网格大小: {grid_size_x}x{grid_size_y} 像素")
    print(f"  - 游戏区域: {cols}列 x {rows}行")
    print(f"  - 起点坐标: ({origin_x}, {origin_y})")

    return grid_info


def grid_to_pixel(grid_x, grid_y, grid_info):
    """将网格坐标转换为屏幕像素坐标。"""
    pixel_x = grid_info["origin_x"] + grid_x * grid_info["size_x"]
    pixel_y = grid_info["origin_y"] + grid_y * grid_info["size_y"]
    return pixel_x, pixel_y


def set_dpi_awareness():
    """设置当前进程为 DPI-aware，以便获取真实分辨率。"""
    # 例如我的 4K 显示器，如果不设置 DPI-aware，由于有默认的150%缩放，GetSystemMetrics 获得的分辨率会是 3840/1.5=2560, 2160/1.5=1440
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        print(
            f"{Colors.OKCYAN}DPI Awareness set to Per-Monitor DPI Aware.{Colors.ENDC}"
        )
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
            print(f"{Colors.OKCYAN}DPI Awareness set to System-DPI Aware.{Colors.ENDC}")
        except (AttributeError, OSError):
            print(
                f"{Colors.WARNING}Could not set DPI awareness (requires Windows Vista or later).{Colors.ENDC}"
            )


def main():
    """主函数。"""
    global game_running

    print(f"{Colors.WARNING}重要提示：此脚本需要以管理员权限运行！{Colors.ENDC}")
    print(
        f"{Colors.WARNING}请随时按 ESC 键退出并恢复桌面。如果用 Ctrl+C 强行退出，会无法恢复桌面。{Colors.ENDC}"
    )

    h_desktop = find_desktop_listview_handle()
    if not h_desktop:
        return

    icon_count = get_icon_count(h_desktop)
    if icon_count < 5:
        print(f"{Colors.FAIL}桌面图标太少({icon_count}个)，无法开始游戏。{Colors.ENDC}")
        return
    print(f"{Colors.OKCYAN}检测到 {icon_count} 个桌面图标。{Colors.ENDC}")

    print(f"\n{Colors.HEADER}请按以下步骤操作，{Colors.ENDC}")
    print(
        "1. 在桌面上，右键 -> 查看 -> 勾选“将图标与网格对齐”。(windows 10/11 下默认已勾选，且无法取消)"
    )
    print("2. 确保有一个图标在屏幕的【左上角】。")
    print("3. 确保有第二个图标在第一个图标【正下方】的紧邻位置。")
    print("4. 确保有第三个图标在第一个图标【正右方】的紧邻位置。")
    print("5. 上述要求是为了计算网格大小。")
    input(f"{Colors.WARNING}完成后请按 Enter 键...{Colors.ENDC}\n")

    # 最小化控制台窗口
    print(f"{Colors.OKCYAN}正在最小化控制台窗口...{Colors.ENDC}")
    time.sleep(0.3)
    console_hwnd = kernel32.GetConsoleWindow()
    if console_hwnd:
        win32gui.ShowWindow(console_hwnd, win32con.SW_MINIMIZE)

    grid_info = calculate_grid_parameters(h_desktop)
    if not grid_info:
        # 在退出前尝试恢复控制台，以便用户看到错误信息
        if console_hwnd:
            win32gui.ShowWindow(console_hwnd, win32con.SW_RESTORE)
        print(f"{Colors.FAIL}初始化失败，程序退出。{Colors.ENDC}")
        return

    keyboard.add_hotkey("esc", stop_game)
    # 最小化后，这些信息在后台打印，用户看不到，但对于调试有用
    print(
        f"\n{Colors.OKGREEN}游戏初始化完成。{Colors.ENDC}"
    )

    try:
        save_initial_positions(h_desktop, icon_count)

        # print("正在准备游戏场地...")

        all_available_icons = list(initial_positions.keys())
        initial_snake_len = 3
        if len(all_available_icons) < initial_snake_len + 2:
            print(
                f"{Colors.FAIL}错误：没有足够的图标来开始游戏（需要至少5个）。{Colors.ENDC}"
            )
            restore_initial_positions(h_desktop)
            return

        snake_indices = deque(all_available_icons[:initial_snake_len])
        waiting_icons = all_available_icons[initial_snake_len:]

        snake_grid_pos = deque([(2, 0), (1, 0), (0, 0)])
        border_grid_pos = []

        # print("正在布置右侧边界...")
        current_col = grid_info["cols"] - 1
        current_row = 0
        for icon_idx in reversed(waiting_icons):
            if current_col < 0:
                print(
                    f"{Colors.WARNING}警告：图标过多，无法在网格内完全展示。{Colors.ENDC}"
                )
                break
            px, py = grid_to_pixel(current_col, current_row, grid_info)
            set_icon_position(h_desktop, icon_idx, px, py)
            border_grid_pos.append((current_col, current_row))
            current_row += 1
            if current_row >= grid_info["rows"]:
                current_row = 0
                current_col -= 1
            time.sleep(0.01)

        # print("正在放置贪吃蛇...")
        for i, icon_idx in enumerate(snake_indices):
            px, py = grid_to_pixel(
                snake_grid_pos[i][0], snake_grid_pos[i][1], grid_info
            )
            set_icon_position(h_desktop, icon_idx, px, py)
            time.sleep(0.01)

        # print("正在放置食物...")
        food_index = waiting_icons.pop(0)
        border_grid_pos.pop()

        border_x_boundary = grid_info["cols"]
        if border_grid_pos:
            border_x_boundary = min(p[0] for p in border_grid_pos)

        max_food_x = border_x_boundary - 3
        if max_food_x < 5:
            print(f"{Colors.FAIL}错误：游戏区域太窄，无法安全放置食物。{Colors.ENDC}")
            restore_initial_positions(h_desktop)
            return

        food_grid_pos = (
            random.randint(5, max_food_x),
            random.randint(1, grid_info["rows"] - 2),
        )
        while food_grid_pos in snake_grid_pos:
            food_grid_pos = (
                random.randint(0, max_food_x),
                random.randint(0, grid_info["rows"] - 1),
            )

        px, py = grid_to_pixel(food_grid_pos[0], food_grid_pos[1], grid_info)
        set_icon_position(h_desktop, food_index, px, py)

        # print("游戏开始！请用 WASD 或方向键控制。")

        direction = (1, 0)
        last_move_time = time.time()

        while game_running:
            # --- 1. 键盘输入 ---
            new_direction = direction
            if keyboard.is_pressed("w") or keyboard.is_pressed("up"):
                if direction != (0, 1):
                    new_direction = (0, -1)
            elif keyboard.is_pressed("s") or keyboard.is_pressed("down"):
                if direction != (0, -1):
                    new_direction = (0, 1)
            elif keyboard.is_pressed("a") or keyboard.is_pressed("left"):
                if direction != (1, 0):
                    new_direction = (-1, 0)
            elif keyboard.is_pressed("d") or keyboard.is_pressed("right"):
                if direction != (-1, 0):
                    new_direction = (1, 0)
            direction = new_direction

            # --- 2. 游戏节拍器 ---
            current_time = time.time()
            base_speed = 0.3
            speed_increase = len(snake_indices) * 0.005
            game_speed = max(0.05, base_speed - speed_increase)

            if current_time - last_move_time >= game_speed:
                last_move_time = current_time

                current_head_pos = snake_grid_pos[0]
                new_head_pos = (
                    current_head_pos[0] + direction[0],
                    current_head_pos[1] + direction[1],
                )

                game_over_message = ""
                if not (
                    0 <= new_head_pos[0] < grid_info["cols"]
                    and 0 <= new_head_pos[1] < grid_info["rows"]
                ):
                    game_over_message = "游戏结束：撞到墙了！"
                elif new_head_pos in snake_grid_pos:
                    game_over_message = "游戏结束：撞到自己了！"
                elif new_head_pos in border_grid_pos:
                    game_over_message = "游戏结束：撞到边界图标了！"

                if game_over_message:
                    print(f"{Colors.FAIL}{game_over_message}{Colors.ENDC}")
                    game_running = False
                    continue

                if new_head_pos == food_grid_pos:
                    snake_grid_pos.appendleft(new_head_pos)
                    snake_indices.appendleft(food_index)

                    if not waiting_icons:
                        print(
                            f"{Colors.BOLD}{Colors.OKGREEN}恭喜你吃完了所有图标，游戏胜利！{Colors.ENDC}"
                        )
                        game_running = False
                        continue

                    food_index = waiting_icons.pop(0)
                    if border_grid_pos:
                        border_grid_pos.pop()

                    border_x_boundary = grid_info["cols"]
                    if border_grid_pos:
                        border_x_boundary = min(p[0] for p in border_grid_pos)
                    max_food_x = max(0, border_x_boundary - 3)

                    while True:
                        food_grid_pos = (
                            random.randint(0, max_food_x),
                            random.randint(0, grid_info["rows"] - 1),
                        )
                        if (
                            food_grid_pos not in snake_grid_pos
                            and food_grid_pos not in border_grid_pos
                        ):
                            break

                    px, py = grid_to_pixel(
                        food_grid_pos[0], food_grid_pos[1], grid_info
                    )
                    set_icon_position(h_desktop, food_index, px, py)

                else:
                    snake_grid_pos.pop()
                    tail_icon_index = snake_indices.pop()

                    snake_grid_pos.appendleft(new_head_pos)
                    snake_indices.appendleft(tail_icon_index)

                    px, py = grid_to_pixel(new_head_pos[0], new_head_pos[1], grid_info)
                    set_icon_position(h_desktop, tail_icon_index, px, py)

            time.sleep(0.01)

    except Exception as e:
        if "console_hwnd" in locals() and console_hwnd:
            win32gui.ShowWindow(console_hwnd, win32con.SW_RESTORE)
        print(f"{Colors.FAIL}游戏主循环发生错误: {e}{Colors.ENDC}")
    finally:
        if "console_hwnd" in locals() and console_hwnd:
            win32gui.ShowWindow(console_hwnd, win32con.SW_RESTORE)
        print(f"{Colors.WARNING}游戏结束，正在恢复桌面...{Colors.ENDC}")
        time.sleep(2)
        restore_initial_positions(h_desktop)
        print(f"{Colors.OKGREEN}桌面已恢复。{Colors.ENDC}")


if __name__ == "__main__":
    set_dpi_awareness()
    main()
