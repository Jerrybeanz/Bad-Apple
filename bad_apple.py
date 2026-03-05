import sys
import os
import xlwings as xw
import cv2
import numpy as np


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        # When executing .exe, PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    else:
        # When running as a script, use current directory
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class VideoPlayer:
    def __init__(self, sheet, length, height):
        self.sheet = sheet
        self.length = length # How many columns to display on Excel sheet
        self.height = height # How many rows to display on Excel sheet
        self.frames = []

    def resize_playback_range(self, column_width, row_height):
        for col in range(1, self.length + 1):
            self.sheet.api.Columns(col).ColumnWidth = column_width

        for row in range(1, self.height + 1):
            self.sheet.api.Rows(row).RowHeight = row_height

    def format_playback_range(self):
        """
        Apply conditional formatting to format cells with value 0 as white and 1 as black
        :return:
        """
        playback_range = self.sheet.range((1, 1), (self.height, self.length))

        # Format 0 as white
        cf_white = playback_range.api.FormatConditions.Add(
            Type=xw.constants.FormatConditionType.xlExpression,
            Formula1="=A1=0"
        )
        cf_white.Interior.Color = xw.utils.rgb_to_int((255, 255, 255))
        cf_white.Font.Color = xw.utils.rgb_to_int((255, 255, 255))

        # Format 1 as black
        cf_black = playback_range.api.FormatConditions.Add(
            Type=xw.constants.FormatConditionType.xlExpression,
            Formula1="=A1=1"
        )
        cf_black.Interior.Color = xw.utils.rgb_to_int((0, 0, 0))
        cf_black.Font.Color = xw.utils.rgb_to_int((0, 0, 0))

    def load_video(self, video_path):
        cap = cv2.VideoCapture(video_path)

        while True:
            ret, frame = cap.read()

            if not ret or frame is None:
                print("End of video")
                break

            resized = cv2.resize(frame, (self.length, self.height), interpolation=cv2.INTER_CUBIC)

            gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)

            _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            # print(binary)

            # cv2.imshow("Bad Apple", binary)
            processed_frame = (binary == 255).astype(np.uint8) # Convert to array of 0 and 1's
            self.frames.append(processed_frame)

    def play_video(self):
        for frame in self.frames:
            self.sheet.range(1, 1).value = frame