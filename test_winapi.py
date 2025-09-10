import win32gui
import win32api
import win32con
import win32process
import commctrl
import ctypes #  提供与C语言兼容的数据类型和函数调用
from ctypes import wintypes

# 定义一些需要的Win32函数和结构体，因为pywin32没有完全封装它们
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

# 定义API函数的参数和返回类型
kernel32.VirtualAllocEx.restype = wintypes.LPVOID #  VirtualAllocEx函数的返回值类型设置为LPVOID（指向void的指针）
kernel32.VirtualAllocEx.argtypes = (wintypes.HANDLE, wintypes.LPVOID, ctypes.c_size_t, wintypes.DWORD, wintypes.DWORD) #  VirtualAllocEx函数的参数类型设置：进程句柄、分配基址、大小、分配类型、保护属性
kernel32.VirtualFreeEx.argtypes = (wintypes.HANDLE, wintypes.LPVOID, ctypes.c_size_t, wintypes.DWORD) #  VirtualFreeEx函数的参数类型设置：进程句柄、要释放的内存地址、大小、释放类型
kernel32.ReadProcessMemory.argtypes = (wintypes.HANDLE, wintypes.LPCVOID, wintypes.LPVOID, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)) #  ReadProcessMemory函数的参数类型设置：进程句柄、要读取的内存地址、存储数据的缓冲区、要读取的字节数、实际读取字节数的指针

# 定义POINT结构体
class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG),
                ("y", wintypes.LONG)]

def get_icon_position(h_listview, index):
    """获取指定索引的图标的位置（使用跨进程内存方法）。"""
    if not h_listview:
        return None, None

    # 1. 获取目标进程句柄
    tid, pid = win32process.GetWindowThreadProcessId(h_listview) #  获取explorer的线程ID和进程ID
    h_process = kernel32.OpenProcess(
        win32con.PROCESS_VM_OPERATION | win32con.PROCESS_VM_READ | win32con.PROCESS_VM_WRITE, #  请求进程访问权限（内存操作、读取、写入）
        False,
        pid
    )
    if not h_process:
        print(f"打开进程失败: {ctypes.get_last_error()}")
        return None, None

    try:
        # 2. 在远程进程中分配内存
        p_remote_buffer = kernel32.VirtualAllocEx( #  在explorer进程中分配内存
            h_process,
            0, #  让系统决定内存分配地址
            ctypes.sizeof(POINT), #  分配的内存大小（POINT结构体的大小）
            win32con.MEM_COMMIT | win32con.MEM_RESERVE, #  内存分配标志：提交和保留内存
            win32con.PAGE_READWRITE #  内存页权限可读可写
        )
        if not p_remote_buffer:
            print(f"远程内存分配失败: {ctypes.get_last_error()}")
            return None, None

        # 3. 发送消息，lParam是远程内存的地址
        result = win32gui.SendMessage(h_listview, commctrl.LVM_GETITEMPOSITION, index, p_remote_buffer)
        
        if result == 0: # 返回值为布尔型，0代表失败
            print(f"SendMessage LVM_GETITEMPOSITION for icon {index} failed.")
            return None, None

        # 4. 从远程进程读取数据
        local_point = POINT()
        bytes_read = ctypes.c_size_t(0)
        
        if not kernel32.ReadProcessMemory(
            h_process,
            p_remote_buffer,
            ctypes.byref(local_point),
            ctypes.sizeof(local_point),
            ctypes.byref(bytes_read)
        ):
            print(f"读取远程内存失败: {ctypes.get_last_error()}")
            return None, None
            
        return local_point.x, local_point.y

    finally:
        # 5. 释放远程内存和进程句柄
        if 'p_remote_buffer' in locals() and p_remote_buffer:
            kernel32.VirtualFreeEx(h_process, p_remote_buffer, 0, win32con.MEM_RELEASE)
        if h_process:
            kernel32.CloseHandle(h_process)


# --- 测试代码 ---
def test_positions():
    h_progman = win32gui.FindWindow("Progman", None)
    h_shelldll = win32gui.FindWindowEx(h_progman, 0, "SHELLDLL_DefView", None)
    h_listview = win32gui.FindWindowEx(h_shelldll, 0, "SysListView32", "FolderView")
    
    icon_count = win32gui.SendMessage(h_listview, commctrl.LVM_GETITEMCOUNT, 0, 0)
    print(f"找到 {icon_count} 个图标。")

    for i in range(icon_count):
        pos = get_icon_position(h_listview, i)
        print(f"图标 {i} 的位置是: {pos}")

if __name__ == "__main__":
    test_positions()