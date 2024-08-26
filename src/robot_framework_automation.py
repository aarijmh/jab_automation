import ctypes
from ctypes import wintypes
import logging
import queue
import threading
from JABWrapper.jab_wrapper import JavaAccessBridgeWrapper
from JABWrapper.context_tree import ContextNode, ContextTree, SearchElement
import time

GetMessage = ctypes.windll.user32.GetMessageW
TranslateMessage = ctypes.windll.user32.TranslateMessage
DispatchMessage = ctypes.windll.user32.DispatchMessageW

def pump_background(pipe: queue.Queue):
    try:
        jab_wrapper = JavaAccessBridgeWrapper()
        pipe.put(jab_wrapper)
        message = ctypes.byref(wintypes.MSG())
        while GetMessage(message, 0, 0, 0) > 0:
            TranslateMessage(message)
            logging.debug("Dispatching msg={}".format(repr(message)))
            DispatchMessage(message)
    except Exception as err:
        print(err)
        pipe.put(None)

def main():
    pipe = queue.Queue()
    thread = threading.Thread(target=pump_background, daemon=True, args=[pipe])
    thread.start()
    jab_wrapper = pipe.get()
    if not jab_wrapper:
        raise Exception("Failed to initialize Java Access Bridge Wrapper")
    time.sleep(0.1) # Wait until the initial messages are parsed, before accessing frames
    print(jab_wrapper)
    win = jab_wrapper.get_windows()
    jab_wrapper.switch_window_by_pid("PopupMessageWindow")
    context_tree = ContextTree(jab_wrapper)
    print(context_tree)
if __name__ == "__main__":
    main()