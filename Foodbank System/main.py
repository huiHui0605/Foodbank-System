import os,sys,json,re,uuid,webbrowser,hashlib
from pathlib import Path
from urllib.parse import quote
from datetime import datetime,date
from openpyxl import Workbook,load_workbook
from PySide6.QtCore import Qt,QTimer
from PySide6.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QPushButton,QLineEdit,QTableWidget,QTableWidgetItem,QHeaderView,QMessageBox,QComboBox,QLabel,QFileDialog,QInputDialog)
from services import generate,ROOT,OUT
SETTINGS=ROOT/'local_settings.json'
def settings():
    try:return json.loads(SETTINGS.read_text(encoding='utf-8'))
    except:return {}
def update_settings(**kw):
    d=settings();d.update(kw);SETTINGS.write_text(json.dumps(d,indent=2),encoding='utf-8')
def database():
    p=Path(settings().get('excel_path',str(ROOT/'records.xlsx')))
    if not p.exists():
        w=Workbook();w.active.title='Records';w.active.append(['Name','Type','Phone','Email','Date']);w.save(p);w.close()
    return p
def cellstr(v):return '' if v is None else (v.isoformat() if isinstance(v,(date,datetime)) else str(v))
def norm(v):return re.sub(r'[^a-z0-9]','',str(v or '').lower())
def identify(headers):
    aliases={'name':['name','fullname','participantname','donorname','volunteername','participant name','nama','nama penuh'],'type':['type','role','category','registrationtype'],'phone':['phone','phonenumber','mobile','contactnumber','whatsapp'],'date':['date','timestamp','submissiondate']}
    return {k:next((i for i,h in enumerate(headers) if norm(h) in [norm(x) for x in names]),None) for k,names in aliases.items()}
class Window(QMainWindow):
    def __init__(self):
        super().__init__();self.setWindowTitle('Community Records | Admin Desktop');self.resize(1400,800)
        self.path=None;self.sheet=None;self.headers=[];self.data=[];self.visible=[];self.dirty=False;self.mtime=0
        root=QWidget();self.setCentralWidget(root);layout=QVBoxLayout(root)
        heading=QLabel('Community Records Manager');heading.setStyleSheet('font-size:22px;font-weight:700;padding:8px');layout.addWidget(heading)
        top=QHBoxLayout();layout.addLayout(top)
        self.search=QLineEdit();self.search.setPlaceholderText('Search any cell...');self.search.textChanged.connect(self.search_changed);top.addWidget(self.search,2)
        self.sheetbox=QComboBox();self.sheetbox.currentTextChanged.connect(self.change_sheet);top.addWidget(self.sheetbox,1)
        for title,fn in [('Import Excel',self.import_excel),('Add Row',self.add_row),('Delete Selected',self.delete_rows),('Save Changes',self.save),('Reload Excel',self.reload),('Certificate Settings',self.configure),('Generate PDFs',self.certificates),('Sync Google Sheets',self.google_sync)]:
            b=QPushButton(title);b.clicked.connect(fn);top.addWidget(b)
        self.table=QTableWidget();self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows);self.table.setAlternatingRowColors(True);self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive);layout.addWidget(self.table)
        self.info=QLabel();layout.addWidget(self.info)
        self.setStyleSheet("""
            QMainWindow, QWidget { background-color: #f4f7fb; color: #172b4d; }
            QLabel { color: #172b4d; background-color: transparent; }
            QLineEdit, QComboBox {
                background-color: #ffffff; color: #172b4d;
                border: 1px solid #b8c6d6; border-radius: 5px;
                padding: 7px; selection-background-color: #2474a6;
                selection-color: #ffffff;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff; color: #172b4d;
                selection-background-color: #d8eaf8; selection-color: #172b4d;
            }
            QPushButton {
                background-color: #12678e; color: #ffffff;
                border: none; border-radius: 5px; padding: 8px;
            }
            QPushButton:hover { background-color: #095475; color: #ffffff; }
            QPushButton:pressed { background-color: #073b55; color: #ffffff; }
            QPushButton:disabled { background-color: #b8c6d6; color: #334155; }
            QTableWidget {
                background-color: #ffffff; color: #172b4d;
                alternate-background-color: #edf4fa;
                gridline-color: #d7e0ea;
                selection-background-color: #c3e2f6;
                selection-color: #10263d;
                border: 1px solid #d7e0ea;
            }
            QTableWidget::item { color: #172b4d; padding: 4px; }
            QTableWidget::item:selected {
                background-color: #c3e2f6; color: #10263d;
            }
            QHeaderView::section {
                background-color: #263b53; color: #ffffff;
                padding: 7px; border: 1px solid #40556e;
                font-weight: 600;
            }
            QTableCornerButton::section { background-color: #263b53; }
            QTableWidget QTableCornerButton { background-color: #263b53; }
            QScrollBar { background-color: #f4f7fb; }
            QToolTip { background-color: #ffffff; color: #172b4d; }
        """)
        self.load_workbook()
        self.timer=QTimer(self);self.timer.timeout.connect(self.watch);self.timer.start(4000)
    def warn(self,title,e):QMessageBox.warning(self,title,str(e))
    def load_workbook(self,requested=None):
        try:
            self.path=database();w=load_workbook(self.path,read_only=True,data_only=False)
            names=w.sheetnames;selected=requested if requested in names else settings().get('sheet',names[0]);selected=selected if selected in names else names[0]
            self.sheetbox.blockSignals(True);self.sheetbox.clear();self.sheetbox.addItems(names);self.sheetbox.setCurrentText(selected);self.sheetbox.blockSignals(False)
            self.sheet=selected;s=w[selected]
            matrix=list(s.values);width=max((len(r) for r in matrix),default=0)
            self.headers=[cellstr(v) or f'Column {i+1}' for i,v in enumerate((matrix[0] if matrix else ())+ (None,)*max(0,width-(len(matrix[0]) if matrix else 0)))]
            self.data=[[cellstr(row[i]) if i<len(row) else '' for i in range(width)] for row in matrix[1:]]
            w.close();self.mtime=self.path.stat().st_mtime_ns;self.dirty=False;self.render()
            self.info.setText(f'{self.path} | Worksheet: {selected} | {len(self.data)} rows, {len(self.headers)} columns')
        except Exception as e:self.warn('Excel load error',e)
    def import_excel(self):
        path,_=QFileDialog.getOpenFileName(self,'Choose an Excel workbook',str(ROOT),'Excel files (*.xlsx *.xlsm)')
        if not path:return
        try:
            if Path(path).name.startswith('~$'):raise ValueError('Please close the workbook in Microsoft Excel first.')
            w=load_workbook(path,read_only=True);names=w.sheetnames;w.close()
            if not names:raise ValueError('The workbook has no worksheets.')
            if QMessageBox.question(self,'Use this workbook',f'Use this Excel as the database?\n{path}\n\nChanges will be saved to the original file. Please make a backup first.')!=QMessageBox.StandardButton.Yes:return
            update_settings(excel_path=str(Path(path).resolve()),sheet=names[0]);self.load_workbook(names[0])
        except Exception as e:self.warn('Import failed',e)
    def change_sheet(self,name):
        if not name or name==self.sheet:return
        if not self.confirm_discard():
            self.sheetbox.blockSignals(True);self.sheetbox.setCurrentText(self.sheet);self.sheetbox.blockSignals(False);return
        update_settings(sheet=name);self.load_workbook(name)
    def confirm_discard(self):
        return not self.dirty or QMessageBox.question(self,'Unsaved changes','Discard unsaved changes?')==QMessageBox.StandardButton.Yes
    def collect(self):
        if not self.dirty:return
        for rownum,index in enumerate(self.visible):
            for c in range(len(self.headers)):
                item=self.table.item(rownum,c)
                if item is not None:self.data[index][c]=item.text()
    def search_changed(self,*args):
        self.collect();self.render()
    def render(self,*args):
        q=self.search.text().casefold();self.visible=[i for i,row in enumerate(self.data) if q in ' '.join(row).casefold()]
        self.table.blockSignals(True);self.table.setRowCount(len(self.visible));self.table.setColumnCount(len(self.headers)+1);self.table.setHorizontalHeaderLabels(self.headers+['Share'])
        for ui,i in enumerate(self.visible):
            for c,v in enumerate(self.data[i]):self.table.setItem(ui,c,QTableWidgetItem(v))
            b=QPushButton('WhatsApp');b.clicked.connect(lambda checked=False,index=i:self.share(index));self.table.setCellWidget(ui,len(self.headers),b)
        self.table.blockSignals(False)
        try:self.table.itemChanged.disconnect(self.edited)
        except (RuntimeError,TypeError):pass
        self.table.itemChanged.connect(self.edited)
    def edited(self,item):self.dirty=True
    def add_row(self):
        self.collect();self.data.append(['']*len(self.headers));self.dirty=True;self.search.clear();self.render();self.table.scrollToBottom()
    def delete_rows(self):
        self.collect();indexes={self.visible[i.row()] for i in self.table.selectedIndexes() if i.row()<len(self.visible)}
        if not indexes:return self.warn('Delete rows','Select at least one row.')
        if QMessageBox.question(self,'Confirm deletion',f'Delete {len(indexes)} row(s) from the original Excel file when saved?')!=QMessageBox.StandardButton.Yes:return
        self.data=[r for i,r in enumerate(self.data) if i not in indexes];self.dirty=True;self.render()
    def save(self):
        try:
            self.collect();w=load_workbook(self.path,keep_vba=self.path.suffix.lower()=='.xlsm');s=w[self.sheet]
            # Only replace the data area in the selected worksheet. Other worksheets remain intact.
            old=max(s.max_row-1,0);new=len(self.data)
            for i,row in enumerate(self.data,2):
                for j,v in enumerate(row,1):s.cell(i,j).value=v if v!='' else None
            if old>new:s.delete_rows(new+2,old-new)
            w.save(self.path);w.close();self.load_workbook(self.sheet);self.info.setText('Saved successfully to the original Excel workbook.')
        except PermissionError:self.warn('Save failed','Close the workbook in Microsoft Excel and try again.')
        except Exception as e:self.warn('Save failed',e)
    def reload(self):
        if self.confirm_discard():self.load_workbook(self.sheet)
    def watch(self):
        if self.path and self.path.exists() and self.path.stat().st_mtime_ns!=self.mtime:
            if self.dirty:self.info.setText('Excel changed externally. Save or reload to resolve changes.')
            else:self.load_workbook(self.sheet)
    def configure(self):
        s=settings();mapping=identify(self.headers)
        for key,label in [('name','Name column'),('type','Certificate type column (donor/volunteer)'),('phone','WhatsApp phone column'),('date','Date column')]:
            opts=['(Not configured)']+self.headers;current=s.get('mapping',{}).get(self.sheet,{}).get(key)
            guess=self.headers[mapping[key]] if mapping.get(key) is not None else None
            selected=current if current in opts else guess
            value,ok=QInputDialog.getItem(self,'Certificate Settings',label,opts,opts.index(selected) if selected in opts else 0,False)
            if not ok:return
            s.setdefault('mapping',{}).setdefault(self.sheet,{})[key]='' if value==opts[0] else value
        SETTINGS.write_text(json.dumps(s,indent=2),encoding='utf-8');self.info.setText('Column mapping saved.')
    def get_field(self,row,key):
        mapping=settings().get('mapping',{}).get(self.sheet,{})
        col=mapping.get(key)
        if col not in self.headers:
            guess=identify(self.headers).get(key);col=self.headers[guess] if guess is not None else None
        return row[self.headers.index(col)] if col in self.headers else ''
    def share(self,i):
        row=self.data[i];name=self.get_field(row,'name') or 'there';phone=re.sub(r'\D','',self.get_field(row,'phone'))
        if phone.startswith('0'):phone='60'+phone[1:]
        message=f'Hello {name}, thank you for your participation in our community project.'
        webbrowser.open(('https://wa.me/'+phone if phone else 'https://api.whatsapp.com/send')+'?text='+quote(message))
        self.info.setText('WhatsApp opened. Send manually and attach any PDF yourself.')
    def certificates(self):
        self.collect();count=0;issues=[]
        for i,row in enumerate(self.data):
            kind=self.get_field(row,'type').strip().lower()
            if kind not in ('donor','volunteer'):continue
            name=self.get_field(row,'name').strip()
            if not name:issues.append(f'Row {i+2}: missing name');continue
            d={'type':kind,'name':name,'date':self.get_field(row,'date'),'external_id':'','record_id':hashlib.sha256(f'{self.sheet}:{i}:{name}'.encode()).hexdigest()[:16],'certificate_signature':'','certificate_pdf':'','status':''}
            try:generate(d);count+=bool(d.get('certificate_pdf'))
            except Exception as e:issues.append(f'Row {i+2}: {e}')
        self.info.setText(f'Generated {count} PDF(s). '+('; '.join(issues[:2]) if issues else ''))
        QMessageBox.information(self,'Certificates',f'{count} PDF(s) generated.\n'+('\n'.join(issues[:8]) if issues else 'Open the certificates folder to view them.'))
    def google_sync(self):
        QMessageBox.information(self,'Google Sheets','Google Sheets syncing is available in the previous structured-records version. This dynamic Excel version does not yet support safe field-by-field synchronization into arbitrary worksheets. No data was changed.')
if __name__=='__main__':
    app=QApplication(sys.argv);w=Window();w.show();sys.exit(app.exec())
