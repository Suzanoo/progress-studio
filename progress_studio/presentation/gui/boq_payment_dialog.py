"""BP-2 review/export dialog. Workers never access Tk variables or widgets."""
from pathlib import Path
from progress_studio.presentation.gui.strings import tr
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.scrolledtext import ScrolledText
from progress_studio.services.boq_payment_export import BOQPaymentExportService


class BOQPaymentDialog(tk.Toplevel):
    def __init__(self,parent,service=None):
        super().__init__(parent)
        self.title(tr('bp2.0'));self.geometry('950x650');self.transient(parent.winfo_toplevel())
        self.service=service or BOQPaymentExportService();self.review=None;self.busy=False
        self.events=queue.Queue();self.cancel=threading.Event();self.controls=[]
        self.variables=[tk.StringVar(self) for _ in range(3)];self.ack=tk.BooleanVar(self,False)
        self.status=tk.StringVar(self,tr('bp2.1'))
        body=ttk.Frame(self,padding=16);body.pack(fill='both',expand=True);body.columnconfigure(1,weight=1)
        for i,label in enumerate((tr('bp2.2'),tr('bp2.3'),tr('bp2.4'))):
            ttk.Label(body,text=label).grid(row=i,column=0,sticky='w')
            entry=ttk.Entry(body,textvariable=self.variables[i]);entry.grid(row=i,column=1,sticky='ew');self.controls.append(entry)
            button=ttk.Button(body,text='Browse...',command=lambda i=i:self.pick(i));button.grid(row=i,column=2);self.controls.append(button)
            self.variables[i].trace_add('write',lambda *_:self.invalidate())
        self.ack_button=ttk.Checkbutton(body,text=tr('bp2.5'),variable=self.ack,command=self.sync)
        self.ack_button.grid(row=3,column=0,columnspan=3,sticky='w')
        self.review_button=ttk.Button(body,text=tr('bp2.6'),command=self.assess);self.review_button.grid(row=4,column=0,columnspan=2,sticky='w')
        self.export_button=ttk.Button(body,text=tr('bp2.7'),command=self.export);self.export_button.grid(row=4,column=2)
        self.text=ScrolledText(body,height=20,wrap='word',state='disabled');self.text.grid(row=5,column=0,columnspan=3,sticky='nsew');body.rowconfigure(5,weight=1)
        ttk.Label(body,textvariable=self.status,wraplength=850).grid(row=6,column=0,columnspan=3,sticky='w')
        self.progress=ttk.Progressbar(body,mode='indeterminate');self.progress.grid(row=7,column=0,columnspan=3,sticky='ew')
        self.close_button=ttk.Button(body,text='Close',command=self.close);self.close_button.grid(row=8,column=2)
        self.protocol('WM_DELETE_WINDOW',self.close);self.sync()

    def invalidate(self):
        self.review=None;self.ack.set(False)
        if hasattr(self,'export_button'):self.sync()

    def sync(self):
        for control in self.controls:control.configure(state='disabled' if self.busy else 'normal')
        self.review_button.configure(state='normal' if not self.busy and all(v.get().strip() for v in self.variables[:2]) else 'disabled')
        ready=self.review is not None and (not self.review.revision_required or self.ack.get())
        self.export_button.configure(state='normal' if not self.busy and ready else 'disabled')
        self.ack_button.configure(state='normal' if not self.busy and self.review is not None and self.review.revision_required else 'disabled')
        self.close_button.configure(text=tr('bp2.14') if self.busy else 'Close')

    def pick(self,index):
        path=filedialog.askopenfilename(parent=self,filetypes=[('Project Session','*.progressstudio')] if index==0 else [('Excel','*.xlsx')])
        if path:self.variables[index].set(path)

    def assess(self):
        if self.busy:return
        paths=tuple(v.get().strip() for v in self.variables);self.invalidate();self.start('review',paths)

    def export(self):
        if self.busy or self.review is None:return
        review=self.review;ack=self.ack.get()
        if review.revision_required and not ack:return
        path=filedialog.asksaveasfilename(parent=self,defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')],initialfile='BOQ_Payment.xlsx')
        if path and self.review is review and not self.busy:self.start('export',(review,Path(path),ack))

    def start(self,action,payload):
        self.busy=True;self.cancel.clear();self.sync();self.progress.start();self.status.set(tr('bp2.8') if action=='review' else tr('bp2.9'))
        service,events,cancelled=self.service,self.events,self.cancel.is_set
        def run():
            try:
                value=service.review(*payload,cancelled=cancelled) if action=='review' else service.export(payload[0],payload[1],acknowledge_revision=payload[2],cancelled=cancelled)
                events.put((action,value,None))
            except Exception as exc:events.put((action,None,exc))
        worker=threading.Thread(target=run,daemon=False)
        try:worker.start()
        except Exception as exc:events.put((action,None,exc))
        self.after(75,self.poll)

    def poll(self):
        try:action,value,error=self.events.get_nowait()
        except queue.Empty:self.after(75,self.poll);return
        self.busy=False;self.progress.stop()
        if error:
            self.review=None;self.status.set(str(error));content=tr('bp2.10')
        elif action=='export':
            self.review=None;self.status.set(f'Exported: {value}');content=tr('bp2.11')
        else:
            self.review=value;self.ack.set(False);self.status.set(tr('bp2.12'))
            p=value.projection
            content=f'{len(p.items)} catalog rows; {len(p.periods)} periods. Projected Requirements, not certified payment.\n\n'
            content+='\n'.join(f'{k}: {v}' for k,v in p.provenance)+'\n\n'
            content+='\n'.join(f'{d.code} | {d.identity} | {d.detail}' for d in p.diagnostics)
        self.text.configure(state='normal');self.text.delete('1.0','end');self.text.insert('1.0',content);self.text.configure(state='disabled');self.sync()

    def close(self):
        if self.busy:self.cancel.set();self.status.set(tr('bp2.13'));return
        self.destroy()
