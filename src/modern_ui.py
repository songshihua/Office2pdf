from pathlib import Path
from PyQt5.QtCore import Qt,QTimer,QRectF,QPointF,QByteArray,QSize
from PyQt5.QtGui import QFont,QPainter,QColor,QLinearGradient,QPainterPath,QPen,QPixmap,QIcon
from PyQt5.QtSvg import QSvgRenderer
from PyQt5.QtWidgets import (QApplication,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,
    QLineEdit,QCheckBox,QPlainTextEdit,QProgressBar,QFrame,QStackedWidget,QScrollArea,QSizePolicy)
from word_pdf_app import Window as LegacyWindow,DEFAULT_DIR

def icon(kind,color='#1764ef'):
    paths={
        'file':'<path d="M6 3h8l5 5v13H6z M14 3v6h5 M9 13h7 M9 17h7"/>',
        'folder':'<path d="M3 7V5h7l2 3h9v13H3z M3 10h18"/>',
        'list':'<path d="M9 6h12 M9 12h12 M9 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>',
        'refresh':'<path d="M21 8a9 9 0 0 0-15-3L3 8 M3 3v5h5 M3 16a9 9 0 0 0 15 3l3-3 M21 21v-5h-5"/>',
        'open':'<path d="M14 3h7v7 M21 3L10 14 M10 5H4v16h16v-6"/>',
    }
    svg=f'<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><g fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{paths[kind]}</g></svg>'
    pix=QPixmap(144,144);pix.fill(Qt.transparent);p=QPainter(pix)
    QSvgRenderer(QByteArray(svg.encode())).render(p);p.end();return QIcon(pix)

def document(p,x,y,w,h,color,label):
    p.save();p.translate(x,y);p.scale(w/66,h/82)
    paper=QLinearGradient(0,0,66,82);paper.setColorAt(0,QColor('white'));paper.setColorAt(1,QColor('#eaf1ff'))
    p.setPen(QPen(QColor('#b6cdf5'),1.2));p.setBrush(paper)
    shape=QPainterPath();shape.moveTo(5,1);shape.lineTo(44,1);shape.lineTo(65,22);shape.lineTo(65,76)
    shape.quadTo(65,81,60,81);shape.lineTo(5,81);shape.quadTo(1,81,1,76);shape.lineTo(1,6);shape.quadTo(1,1,5,1)
    p.drawPath(shape)
    fold=QPainterPath();fold.moveTo(44,1);fold.lineTo(44,22);fold.lineTo(65,22)
    p.setBrush(QColor('#d5e4ff'));p.drawPath(fold)
    p.setPen(QPen(QColor('#b2ccf7'),3.2,Qt.SolidLine,Qt.RoundCap))
    for yy,ll in [(39,17),(50,27),(61,20)]:p.drawLine(QPointF(29,yy),QPointF(29+ll,yy))
    if label:
        gradient=QLinearGradient(0,24,40,60);gradient.setColorAt(0,QColor(color).lighter(112));gradient.setColorAt(1,QColor(color))
        p.setPen(Qt.NoPen);p.setBrush(gradient);p.drawRoundedRect(QRectF(-10,26,42,35),5,5)
        f=QFont('Microsoft YaHei');f.setPixelSize(22 if len(label)==1 else 15);f.setBold(True);p.setFont(f);p.setPen(Qt.white)
        p.drawText(QRectF(-10,26,42,35),Qt.AlignCenter,label)
    p.restore()

class DocumentArt(QWidget):
    def __init__(self,empty=False):
        super().__init__();self.empty=empty;self.setSizePolicy(QSizePolicy.Preferred,QSizePolicy.Preferred)
    def paintEvent(self,event):
        p=QPainter(self);p.setRenderHint(QPainter.Antialiasing)
        if not self.empty:
            scale=min(self.width()/240,self.height()/100);p.translate((self.width()-240*scale)/2,(self.height()-100*scale)/2);p.scale(scale,scale)
            document(p,20,7,63,82,'#1764ef','W');document(p,165,7,63,82,'#f14547','PDF')
            p.setPen(QPen(QColor('#2876f5'),4.5,Qt.SolidLine,Qt.RoundCap,Qt.RoundJoin))
            p.drawLine(QPointF(110,48),QPointF(143,48));p.drawLine(QPointF(134,39),QPointF(143,48));p.drawLine(QPointF(134,57),QPointF(143,48))
        else:
            scale=min(self.width()/220,self.height()/125);p.translate((self.width()-220*scale)/2,(self.height()-125*scale)/2);p.scale(scale,scale)
            bg=QPainterPath();bg.moveTo(2,113);bg.cubicTo(45,61,64,99,89,86);bg.cubicTo(136,60,140,51,174,88);bg.cubicTo(198,83,204,105,219,113)
            g=QLinearGradient(0,65,0,120);g.setColorAt(0,QColor('#dceaff'));g.setColorAt(1,QColor('#ffffff'))
            p.setBrush(g);p.setPen(Qt.NoPen);p.drawPath(bg);document(p,77,6,67,86,'#4b8fff','')
            p.setBrush(QColor('#2876ff'));p.setPen(QPen(QColor('#1764ed'),1));p.drawEllipse(QPointF(145,90),19,19)
            p.setPen(QPen(Qt.white,3.3,Qt.SolidLine,Qt.RoundCap));p.drawLine(QPointF(135,90),QPointF(155,90));p.drawLine(QPointF(145,80),QPointF(145,100))
        p.end()

class Banner(QFrame):
    def paintEvent(self,event):
        p=QPainter(self);p.setRenderHint(QPainter.Antialiasing)
        gradient=QLinearGradient(0,0,self.width(),self.height());gradient.setColorAt(0,QColor('#f0f6ff'));gradient.setColorAt(1,QColor('#e4eeff'))
        p.setPen(Qt.NoPen);p.setBrush(gradient);p.drawRoundedRect(self.rect(),15,15)
        ribbon=QPainterPath();ribbon.moveTo(self.width()*.54,self.height());ribbon.cubicTo(self.width()*.72,-25,self.width()*.81,30,self.width(),-20);ribbon.lineTo(self.width(),18);ribbon.cubicTo(self.width()*.82,54,self.width()*.77,22,self.width()*.64,self.height());ribbon.closeSubpath()
        p.setBrush(QColor(207,225,255,65));p.drawPath(ribbon);p.end()

class IconWidget(QWidget):
    def __init__(self,kind):super().__init__();self.picture=icon(kind)
    def paintEvent(self,event):
        painter=QPainter(self);painter.setRenderHint(QPainter.Antialiasing)
        self.picture.paint(painter,self.rect());painter.end()

class Window(LegacyWindow):
    def __init__(self):
        QWidget.__init__(self);self.worker=None;self.result=None;self._scale=None
        self.setWindowTitle('Word 转 PDF');self.setWindowIcon(icon('file'));self.setAcceptDrops(True);self.setMinimumSize(660,540)
        screen=QApplication.primaryScreen().availableGeometry()
        self.resize(min(1120,int(screen.width()*.93)),min(830,int(screen.height()*.91)))
        outer=QVBoxLayout(self);outer.setContentsMargins(0,0,0,0);outer.setSpacing(0)
        self.scroll=QScrollArea();self.scroll.setFrameShape(QFrame.NoFrame);self.scroll.setWidgetResizable(True);self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff);outer.addWidget(self.scroll)
        self.page=QWidget();self.page.setObjectName('page');self.scroll.setWidget(self.page)
        self.layout_root=QVBoxLayout(self.page);self.layout_root.setSpacing(12)
        self.banner=Banner();self.header_layout=QHBoxLayout(self.banner);words=QVBoxLayout();words.setSpacing(8)
        self.title=QLabel('将 Word 文档转换为 PDF');self.title.setObjectName('title');self.title.setWordWrap(True)
        self.subtitle=QLabel('保留原排版、公式与图片，方便分享和打印。');self.subtitle.setObjectName('muted');self.subtitle.setWordWrap(True)
        words.addWidget(self.title);words.addWidget(self.subtitle);self.header_layout.addLayout(words,1)
        self.hero_art=DocumentArt();self.header_layout.addWidget(self.hero_art);self.layout_root.addWidget(self.banner)
        self.section_icons=[];self.layout_root.addLayout(self.section('file','选择 Word 文件'))
        row=QHBoxLayout();self.source=QLineEdit();self.source.setPlaceholderText('选择 .docx / .doc / .docm 文件');self.source.setToolTip('也可以将 Word 文件直接拖入窗口')
        self.browse=QPushButton('选择文件');self.browse.setIcon(icon('folder','#22334b'));self.browse.clicked.connect(self.pick_source)
        row.addWidget(self.source,1);row.addWidget(self.browse);self.layout_root.addLayout(row)
        self.layout_root.addLayout(self.section('folder','PDF 保存位置'))
        row=QHBoxLayout();self.folder=QLineEdit(str(DEFAULT_DIR));self.pick_folder=QPushButton('更改位置');self.pick_folder.setIcon(icon('folder','#22334b'));self.pick_folder.clicked.connect(self.pick_dir)
        row.addWidget(self.folder,1);row.addWidget(self.pick_folder);self.layout_root.addLayout(row)
        self.compat=QCheckBox('跨设备兼容模式（公式、复杂图表等内容稳定显示）');self.compat.setChecked(True)
        self.compat.setToolTip('保留 SVG 矢量图，处理公式字体与渐变。截图保留原清晰度，原 Word 不会被修改。');self.layout_root.addWidget(self.compat)
        self.convert_btn=QPushButton('转换为 PDF');self.convert_btn.setIcon(icon('refresh','white'));self.convert_btn.setObjectName('convert');self.convert_btn.clicked.connect(self.start);self.layout_root.addWidget(self.convert_btn)
        self.separator=QFrame();self.separator.setObjectName('separator');self.separator.setFixedHeight(1);self.layout_root.addWidget(self.separator);self.layout_root.addLayout(self.section('list','转换进度'))
        self.card=QFrame();self.card.setObjectName('taskCard');card_layout=QVBoxLayout(self.card);self.stack=QStackedWidget();card_layout.addWidget(self.stack)
        self.empty=QWidget();empty_layout=QVBoxLayout(self.empty);empty_layout.setSpacing(7);empty_layout.setContentsMargins(12,6,12,10);empty_layout.addStretch(1)
        self.empty_art=DocumentArt(True);empty_layout.addWidget(self.empty_art,0,Qt.AlignHCenter)
        self.empty_title=QLabel('暂无转换任务');self.empty_title.setObjectName('emptyTitle');self.empty_title.setAlignment(Qt.AlignCenter)
        self.empty_help=QLabel('请先选择 Word 文件并点击“转换为 PDF”');self.empty_help.setObjectName('muted');self.empty_help.setWordWrap(True);self.empty_help.setAlignment(Qt.AlignCenter)
        empty_layout.addWidget(self.empty_title);empty_layout.addWidget(self.empty_help);empty_layout.addStretch(1);self.stack.addWidget(self.empty)
        self.progress_panel=QWidget();progress_layout=QVBoxLayout(self.progress_panel);progress_layout.setContentsMargins(10,8,10,8)
        self.status=QLabel('就绪');self.status.setObjectName('status');self.status.setWordWrap(True);progress_layout.addWidget(self.status)
        self.bar=QProgressBar();self.bar.setRange(0,100);self.bar.setValue(0);self.bar.setTextVisible(False);progress_layout.addWidget(self.bar)
        self.log=QPlainTextEdit();self.log.setObjectName('log');self.log.setReadOnly(True);progress_layout.addWidget(self.log,1);self.stack.addWidget(self.progress_panel);self.layout_root.addWidget(self.card,1)
        footer=QWidget();footer.setObjectName('page');bottom=QHBoxLayout(footer);self.footer_layout=bottom;bottom.addStretch(1)
        self.open_file=QPushButton('打开 PDF');self.open_file.setIcon(icon('open','#627d9f'));self.open_file.setEnabled(False);self.open_file.clicked.connect(self.open_result)
        self.open_folder=QPushButton('打开保存文件夹');self.open_folder.setIcon(icon('folder','#22334b'));self.open_folder.clicked.connect(self.show_folder)
        bottom.addWidget(self.open_file);bottom.addWidget(self.open_folder);outer.addWidget(footer)
        self.resize_timer=QTimer(self);self.resize_timer.setSingleShot(True);self.resize_timer.setInterval(35);self.resize_timer.timeout.connect(self.apply_scale);self.apply_scale()

    def section(self,kind,title):
        row=QHBoxLayout();row.setContentsMargins(0,2,0,0);row.setSpacing(10)
        pic=IconWidget(kind);self.section_icons.append((pic,kind));label=QLabel(title);label.setObjectName('section')
        row.addWidget(pic);row.addWidget(label);row.addStretch();return row
    def resizeEvent(self,event):
        QWidget.resizeEvent(self,event)
        if hasattr(self,'resize_timer'):self.resize_timer.start()
    def apply_scale(self):
        # Qt logical pixels account for DPI; resize typography and controls together.
        scale=max(.84,min(1.85,min(self.width()/1120,self.height()/800)))
        if self._scale is not None and abs(scale-self._scale)<.015:return
        self._scale=scale
        def px(v):return max(1,round(v*scale))
        self.layout_root.setContentsMargins(px(28),px(16),px(28),px(9));self.layout_root.setSpacing(px(9));self.header_layout.setContentsMargins(px(24),px(12),px(22),px(12))
        self.footer_layout.setContentsMargins(px(28),px(3),px(28),px(16))
        self.banner.setMinimumHeight(px(112));self.hero_art.setFixedSize(px(235),px(80));self.empty_art.setFixedSize(px(196),px(94));self.card.setMinimumHeight(px(180))
        for pic,kind in self.section_icons:
            size=px(24);pic.setFixedSize(size,size)
        for button in (self.browse,self.pick_folder,self.open_file,self.open_folder):button.setMinimumHeight(px(47));button.setIconSize(QSize(px(22),px(22)))
        for button in (self.browse,self.pick_folder):button.setMinimumWidth(px(138))
        for edit in (self.source,self.folder):edit.setMinimumHeight(px(47))
        self.compat.setMinimumHeight(px(38));self.convert_btn.setMinimumHeight(px(59));self.convert_btn.setIconSize(QSize(px(27),px(27)))
        self.setStyleSheet(f'''
        QWidget{{font-family:"Microsoft YaHei";font-size:{px(18)}px;color:#162133;}}
        QWidget#page,QScrollArea{{background:#f6f8fd;border:0;}} QLabel{{background:transparent;}}
        QLabel#title{{font-size:{px(30)}px;font-weight:700;color:#0c1524;}}
        QLabel#muted{{color:#68768b;font-size:{px(17)}px;}} QLabel#section{{font-size:{px(19)}px;font-weight:700;}}
        QLabel#emptyTitle{{font-size:{px(23)}px;font-weight:700;color:#526078;}} QLabel#status{{font-size:{px(19)}px;font-weight:600;}}
        QLineEdit{{background:white;border:1px solid #c7d5eb;border-radius:{px(9)}px;padding:{px(5)}px {px(15)}px;selection-background-color:#dbe8ff;}}
        QLineEdit:focus{{border:2px solid #4c89f5;}}
        QPushButton{{background:white;border:1px solid #c7d4e8;border-radius:{px(9)}px;padding:{px(6)}px {px(17)}px;font-size:{px(18)}px;font-weight:600;}}
        QPushButton:hover{{background:#eaf2ff;border-color:#8baff3;}} QPushButton:pressed{{background:#dae8ff;}}
        QPushButton:disabled{{background:#eaf0f8;color:#9aa9be;border-color:#dce5f2;}}
        QPushButton#convert{{color:white;background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #1464ec,stop:1 #0d5ce5);border:0;font-size:{px(23)}px;font-weight:700;}}
        QPushButton#convert:hover{{background:#0755d8;}} QPushButton#convert:disabled{{background:#95b7ee;color:#edf3ff;}}
        QCheckBox{{font-size:{px(18)}px;spacing:{px(12)}px;font-weight:600;}} QCheckBox::indicator{{width:{px(23)}px;height:{px(23)}px;}}
        QFrame#separator{{background:#d9e2ef;}} QFrame#taskCard{{background:white;border:1px solid #ccdaee;border-radius:{px(12)}px;}}
        QStackedWidget,QStackedWidget>QWidget{{background:transparent;}}
        QPlainTextEdit#log{{background:white;border:0;padding:{px(4)}px;font-size:{px(16)}px;color:#485b75;}}
        QProgressBar{{background:#e7eef9;border:0;border-radius:{px(4)}px;min-height:{px(7)}px;max-height:{px(7)}px;}} QProgressBar::chunk{{background:#2472ef;border-radius:{px(4)}px;}}
        QScrollBar:vertical{{background:#eff3fa;width:{px(10)}px;border:0;}} QScrollBar::handle:vertical{{background:#c5d1e4;min-height:30px;border-radius:4px;}}
        ''')
    def start(self):
        source=Path(self.source.text().strip().strip('"'))
        if source.is_file() and source.suffix.lower() in ('.docx','.doc','.docm') and self.folder.text().strip():self.stack.setCurrentWidget(self.progress_panel)
        super().start()
    def set_busy(self,busy):
        super().set_busy(busy);self.convert_btn.setText('正在转换…' if busy else '转换为 PDF')
