
import time
import datetime
import subprocess
import threading
from winsdk.windows.ui.notifications import ToastNotificationManager, ToastNotification
from winsdk.windows.data.xml.dom import XmlDocument
import winsdk.windows.foundation as foundation

# --- 全局变量 ---
# 用于存储用户操作的变量 (例如 'snooze' 或 'ignore')
user_action = None
# 用于在点击通知按钮后发出信号的事件
notification_activated = threading.Event()

# --- 通知处理 ---

def on_toast_activated(sender, args):
    """当用户与通知交互时调用的回调函数。"""
    global user_action
    try:
        # 从通知参数中获取用户操作
        user_action = args.arguments
    except Exception as e:
        print(f"Error getting arguments: {e}")
        user_action = 'ignore' # 出错时默认为忽略
    finally:
        # 设置事件，表示通知已被处理
        notification_activated.set()

def show_shutdown_toast():
    """创建并显示关机提醒通知。"""
    # 定义通知的 XML 结构
    toast_xml = f"""
    <toast launch="shutdown-reminder" scenario="reminder">
        <visual>
            <binding template="ToastGeneric">
                <text>即将关机</text>
                <text>白板将在3分钟后自动关机。</text>
                <text>若未完成，可选择延迟关机以继续授课。</text>
            </binding>
        </visual>
        <actions>
            <action
                content="延迟10分钟"
                arguments="snooze"
                activationType="background"/>
            <action
                content="忽略"
                arguments="ignore"
                activationType="background"/>
        </actions>
    </toast>
    """
    # 创建 XML 文档对象
    xml_doc = XmlDocument()
    xml_doc.load_xml(toast_xml)

    # 创建通知对象
    toast = ToastNotification(xml_doc)

    # 关键：为通知的 'activated' 事件添加回调函数
    # 使用 lambda 确保在 UI 线程上正确调度
    activated_handler = foundation.TypedEventHandler[ToastNotification, object](on_toast_activated)
    toast.add_activated(activated_handler)

    # 显示通知
    ToastNotificationManager.get_default().create_toast_notifier().show(toast)
    print("Push shutdown notification successed.")

# --- 关机逻辑 ---

def shutdown(delay_seconds):
    """
    在指定的延迟后执行关机命令。
    :param delay_seconds: 关机前的等待秒数。
    """
    print(f"计划在 {delay_seconds} 秒后关机...")
    subprocess.run(f"shutdown /s /t {delay_seconds}", shell=True)

def handle_shutdown_logic():
    """处理显示通知和根据用户响应执行关机。"""
    global user_action
    
    show_shutdown_toast()

    # 等待用户交互或超时 (3分钟)
    # notification_activated.wait() 会阻塞直到 on_toast_activated 被调用
    # timeout=180 表示如果180秒内没有交互，就继续执行
    activated = notification_activated.wait(timeout=180)

    if not activated:
        # 如果超时（用户未点击任何按钮）
        print("用户未在3分钟内响应。")
        user_action = 'ignore' # 视为忽略

    if user_action == 'snooze':
        print("用户选择延迟。将在10分钟后关机。")
        shutdown(600)  # 10分钟 = 600秒
    else: # 'ignore' 或其他情况
        print("用户选择忽略或未响应。将在3分钟后关机。")
        # 由于我们已经等待了3分钟，所以立即关机
        # 为了安全起见，可以设置一个很短的延迟
        shutdown(1)

# --- 主程序 ---

def main():
    """主函数，直接执行关机逻辑。"""
    print("自动化脚本已启动，开始执行关机逻辑。")
    # 在新线程中处理关机逻辑，以防主循环被阻塞
    shutdown_thread = threading.Thread(target=handle_shutdown_logic)
    shutdown_thread.start()
    # 等待关机线程结束
    shutdown_thread.join()
    print("脚本执行完毕。")

if __name__ == "__main__":
    main()
