# -*- coding: utf-8 -*-
"""生成服务器版图标（多尺寸 PNG）：蓝色圆角底 + 白色文档 + 服务器机架 + 运行指示点。

输出：app/web/icon.png（256）与 icon_16/32/48/64.png
运行：QT_QPA_PLATFORM=offscreen app/runtime/python.exe tools/make_server_icon.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6.QtCore import QPointF, QRectF, Qt
from PyQt6.QtGui import (QBrush, QColor, QLinearGradient, QPainter, QPixmap,
                         QPolygonF)
from PyQt6.QtWidgets import QApplication

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "app", "web")


def draw(size):
    pm = QPixmap(size, size)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    p.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
    s = float(size)

    # 1) 圆角渐变底
    g = QLinearGradient(0, 0, s * 0.35, s)
    g.setColorAt(0.0, QColor("#4f8ef7"))
    g.setColorAt(1.0, QColor("#1d4ed8"))
    p.setBrush(QBrush(g))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRoundedRect(QRectF(s * 0.04, s * 0.04, s * 0.92, s * 0.92), s * 0.22, s * 0.22)

    # 2) 白色文档（带折角）
    dx, dy, dw, dh = s * 0.20, s * 0.17, s * 0.42, s * 0.56
    fold = s * 0.10
    p.setBrush(QBrush(QColor("#ffffff")))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRoundedRect(QRectF(dx, dy, dw, dh), s * 0.045, s * 0.045)
    # 折角三角（用浅色覆盖右上角）
    p.setBrush(QBrush(QColor("#dbeafe")))
    p.drawPolygon(QPolygonF([
        QPointF(dx + dw - fold, dy),
        QPointF(dx + dw, dy),
        QPointF(dx + dw, dy + fold),
    ]))

    # 3) 文档文字行
    p.setBrush(QBrush(QColor("#94a3b8")))
    line_x = dx + dw * 0.16
    line_w = dw * 0.68
    for i, h in enumerate((0.30, 0.42, 0.54, 0.66)):
        ww = line_w if i < 3 else line_w * 0.55
        p.drawRoundedRect(QRectF(line_x, dy + dh * h, ww, max(1.0, s * 0.028)),
                          s * 0.014, s * 0.014)

    # 4) 服务器机架（右下三格）
    rx, ry, rw = s * 0.56, s * 0.40, s * 0.26
    rh = s * 0.115
    for i in range(3):
        y = ry + i * (rh + s * 0.035)
        p.setBrush(QBrush(QColor("#ffffff")))
        p.drawRoundedRect(QRectF(rx, y, rw, rh), s * 0.025, s * 0.025)
        # 指示灯
        p.setBrush(QBrush(QColor("#22c55e") if i == 0 else QColor("#93c5fd")))
        d = s * 0.028
        p.drawEllipse(QRectF(rx + rw - d - s * 0.022, y + (rh - d) / 2, d, d))
        # 硬盘槽
        p.setBrush(QBrush(QColor("#cbd5e1")))
        p.drawRoundedRect(
            QRectF(rx + s * 0.025, y + (rh - s * 0.02) / 2, rw * 0.45, max(1.0, s * 0.02)),
            s * 0.01, s * 0.01)

    p.end()
    return pm


def main():
    app = QApplication(sys.argv)  # noqa: F841
    os.makedirs(OUT_DIR, exist_ok=True)
    pm = draw(256)
    main_path = os.path.join(OUT_DIR, "icon.png")
    pm.save(main_path, "PNG")
    print("已生成:", main_path)
    for sz in (16, 32, 48, 64):
        p = draw(sz)
        pth = os.path.join(OUT_DIR, "icon_%d.png" % sz)
        p.save(pth, "PNG")
        print("已生成:", pth)


if __name__ == "__main__":
    main()
