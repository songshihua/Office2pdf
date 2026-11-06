import sys
import os
import traceback
import json
from pathlib import Path
from datetime import datetime

from PyQt5.QtCore import Qt,QThread,pyqtSignal,QTimer
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (QApplication,QWidget,QVBoxLayout,QHBoxLayout,QLabel,
    QPushButton,QFileDialog,QLineEdit,QCheckBox,QPlainTextEdit,QProgressBar,QMessageBox)
from converter_engine import convert,register_fonts

DEFAULT_DIR=Path(r'D:\360Downloads\pdf转换')

class Worker(QThread):
    progress=pyqtSignal(str);success=pyqtSignal(dict);failure=pyqtSignal(str)
    def __init__(self,source,folder,compatibility):
        super().__init__();self.source=source;self.folder=folder;self.compatibility=compatibility
    def run(self):
        try:self.success.emit(convert(self.source,self.folder,self.compatibility,self.progress.emit))
        except Exception:
            self.failure.emit(traceback.format_exc())

class Window(QWidget):
    def __init__(self):
        super().__init__();self.worker=None;self.result=None
        self.setWindowTitle('Word 转 PDF · 高清与兼容');self.resize(760,650)
        self.setMinimumSize(680,560);self.setAcceptDrops(True)
        self.setStyleSheet('''QWidget{background:#f5f7fb;color:#182c46;font-family:"Microsoft YaHei";font-size:14px;}
        QLabel#title{font-size:27px;font-weight:700;} QLabel#muted{color:#62748c;}
        QLineEdit,QPlainTextEdit{background:white;border:1px solid #dce3ed;border-radius:8px;padding:10px;}
        QPushButton{background:white;border:1px solid #cad5e5;border-radius:8px;padding:10px 16px;}
        QPushButton:hover{background:#edf3fc;} QPushButton:disabled{color:#9aa9bb;background:#e8edf5;}
        QPushButton#convert{background:#2467db;color:white;border:0;font-size:17px;font-weight:700;padding:15px;}
        QPushButton#convert:hover{background:#1853ba;} QProgressBar{border:0;background:#e2e9f3;border-radius:4px;max-height:8px;}
        QProgressBar::chunk{background:#2467db;border-radius:4px;}''')
        layout=QVBoxLayout(self);layout.setContentsMargins(30,26,30,25);layout.setSpacing(15)
        title=QLabel('把 Word 转成清晰的 PDF');title.setObjectName('title');layout.addWidget(title)
        label=QLabel('保留原排版，恢复原始矢量图，并处理公式的跨设备显示。');label.setObjectName('muted');layout.addWidget(label)
        layout.addWidget(QLabel('Word 文件（也可直接拖入窗口）'))
        row=QHBoxLayout();self.source=QLineEdit();self.source.setPlaceholderText('选择 .docx / .doc / .docm 文件')
        self.browse=QPushButton('选择文件');self.browse.clicked.connect(self.pick_source)
        row.addWidget(self.source,1);row.addWidget(self.browse);layout.addLayout(row)
        layout.addWidget(QLabel('PDF 保存位置'))
        row=QHBoxLayout();self.folder=QLineEdit(str(DEFAULT_DIR));self.pick_folder=QPushButton('更改位置');self.pick_folder.clicked.connect(self.pick_dir)
        row.addWidget(self.folder,1);row.addWidget(self.pick_folder);layout.addLayout(row)
        self.compat=QCheckBox('跨设备兼容模式（公式转矢量轮廓、简化渐变）');self.compat.setChecked(True);layout.addWidget(self.compat)
        hint=QLabel('原有截图保持原清晰度，不自动描摹图表。原 Word 不会被修改；同名结果自动另存。')
        hint.setWordWrap(True);hint.setObjectName('muted');layout.addWidget(hint)
        self.convert_btn=QPushButton('转换为 PDF');self.convert_btn.setObjectName('convert');self.convert_btn.clicked.connect(self.start);layout.addWidget(self.convert_btn)
        self.bar=QProgressBar();self.bar.setRange(0,100);self.bar.setValue(0);self.bar.setTextVisible(False);layout.addWidget(self.bar)
        self.status=QLabel('就绪 · 需要本机安装 WPS 或 Microsoft Word');self.status.setWordWrap(True);layout.addWidget(self.status)
        self.log=QPlainTextEdit();self.log.setReadOnly(True);self.log.setPlaceholderText('转换进度会显示在这里。');layout.addWidget(self.log,1)
        row=QHBoxLayout();row.addStretch();self.open_file=QPushButton('打开 PDF');self.open_file.setEnabled(False);self.open_file.clicked.connect(self.open_result)
        self.open_folder=QPushButton('打开保存文件夹');self.open_folder.clicked.connect(self.show_folder)
        row.addWidget(self.open_file);row.addWidget(self.open_folder);layout.addLayout(row)
    def pick_source(self):
        path,_=QFileDialog.getOpenFileName(self,'选择 Word 文件','','Word 文件 (*.docx *.doc *.docm)')
        if path:self.source.setText(path)
    def pick_dir(self):
        path=QFileDialog.getExistingDirectory(self,'选择 PDF 保存位置',self.folder.text())
        if path:self.folder.setText(path)
    def dragEnterEvent(self,event):
        if event.mimeData().hasUrls() and not(self.worker and self.worker.isRunning()):event.acceptProposedAction()
    def dropEvent(self,event):
        urls=event.mimeData().urls()
        if urls:self.source.setText(urls[0].toLocalFile())
    def start(self):
        source=Path(self.source.text().strip().strip('"'));folder=Path(self.folder.text().strip().strip('"'))
        if not source.is_file() or source.suffix.lower() not in ('.docx','.doc','.docm'):
            QMessageBox.information(self,'请选择文件','请先选择一个存在的 Word 文件。');return
        if not self.folder.text().strip():QMessageBox.information(self,'请选择位置','请先选择 PDF 保存文件夹。');return
        self.result=None;self.log.clear();self.set_busy(True);self.open_file.setEnabled(False)
        self.status.setText('正在转换，请保持文字处理程序可用…')
        self.worker=Worker(str(source),str(folder),self.compat.isChecked())
        self.worker.progress.connect(self.log.appendPlainText);self.worker.success.connect(self.finished)
        self.worker.failure.connect(self.failed);self.worker.finished.connect(lambda:self.set_busy(False));self.worker.start()
    def set_busy(self,busy):
        for widget in (self.convert_btn,self.browse,self.pick_folder,self.source,self.folder,self.compat):widget.setEnabled(not busy)
        self.bar.setRange(0,0 if busy else 100)
        if not busy:self.bar.setValue(100 if self.result else 0)
    def finished(self,result):
        self.result=result;self.open_file.setEnabled(True)
        self.status.setText(f'转换完成 · {result["pages"]} 页 · 已恢复 {result["vectors"]} 幅矢量图')
        self.log.appendPlainText('\n已保存：'+result['path'])
        for warning in result['warnings']:self.log.appendPlainText('提示：'+warning)
        self.save_log()
    def failed(self,error):
        self.status.setText('转换未完成，请查看下方说明。');self.log.appendPlainText(error);self.save_log()
        QMessageBox.warning(self,'转换未完成','没有生成未完成的 PDF。请查看窗口中的错误信息或转换日志。')
    def save_log(self):
        try:
            folder=Path(self.folder.text())/'转换日志';folder.mkdir(parents=True,exist_ok=True)
            name=datetime.now().strftime('%Y%m%d_%H%M%S')+'.txt';(folder/name).write_text(self.log.toPlainText(),encoding='utf-8')
        except OSError:pass
    def open_result(self):
        if self.result:os.startfile(self.result['path'])
    def show_folder(self):
        try:
            folder=Path(self.folder.text());folder.mkdir(parents=True,exist_ok=True);os.startfile(str(folder))
        except OSError as exc:QMessageBox.warning(self,'无法打开',str(exc))
    def closeEvent(self,event):
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self,'正在转换','转换仍在进行，请完成后再关闭窗口。');event.ignore()
        else:event.accept()

def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling,True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps,True)
    if hasattr(Qt,'HighDpiScaleFactorRoundingPolicy'):
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app=QApplication(sys.argv);app.setApplicationName('Word 转 PDF');register_fonts(app)
    if '--convert' in sys.argv:
        try:
            i=sys.argv.index('--convert');source=sys.argv[i+1]
            folder=Path(sys.argv[sys.argv.index('--output-dir')+1]) if '--output-dir' in sys.argv else DEFAULT_DIR
            result=convert(source,folder,'--standard' not in sys.argv,lambda m:print(m,flush=True))
            if '--result-json' in sys.argv:
                Path(sys.argv[sys.argv.index('--result-json')+1]).write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
            return 0
        except Exception:
            if '--result-json' in sys.argv:Path(sys.argv[sys.argv.index('--result-json')+1]).write_text(traceback.format_exc(),encoding='utf-8')
            return 1
    from modern_ui import Window as ModernWindow
    window=ModernWindow()
    if '--preview-size' in sys.argv:
        w,h=map(int,sys.argv[sys.argv.index('--preview-size')+1].split('x'));window.resize(w,h)
    window.show()
    if '--smoke-ui' in sys.argv:
        path=sys.argv[sys.argv.index('--smoke-ui')+1]
        if '--preview-state' in sys.argv and sys.argv[sys.argv.index('--preview-state')+1]=='busy':
            window.stack.setCurrentWidget(window.progress_panel);window.status.setText('正在转换，请稍候…');window.set_busy(True)
            window.log.appendPlainText('正在排版并导出文档…\n正在恢复原始矢量图…')
        def grab():window.grab().save(path);app.quit()
        QTimer.singleShot(800,grab)
    return app.exec_()

if __name__=='__main__':
    sys.exit(main())
