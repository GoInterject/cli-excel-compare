import sys

class Informer:
    
    def __init__(self):
        self.red_text = "\033[31m"
        self.green_text = "\033[32m"
        self.yellow_text = "\033[33m"
        self.default_attrib_text = "\033[0m"

        self.info_level_notification = 1
        self.info_level_warning = 2
        self.info_level_error = 3

        self.info_level_texts = {
            self.info_level_warning: {
                "text": "Warning",
                "color": self.yellow_text
                },
            self.info_level_error: {
                "text": "Error",
                "color": self.red_text
                }
        }


    def set_cursor_x_position(self, x):
        sys.stdout.write(f"\033[{x}G")
        sys.stdout.flush()

    def set_color_text(self, text: str, color: str):
        return f"{color}{text}{self.default_attrib_text}"

    def print_info(self, info_level: str, info_message: str):
        info_level_text = None
        if info_level == self.info_level_notification:
            info_level_text = ""
        else:
            info_level_text: str = self.info_level_texts[info_level]["text"]
            info_level_text = f"{info_level_text} message. "
            info_level_text = self.set_color_text(info_level_text, self.info_level_texts[info_level]["color"])
        print(f"{info_level_text}{info_message}")
