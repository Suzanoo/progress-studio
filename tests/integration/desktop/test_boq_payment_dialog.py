"""Headless orchestration proof; not native desktop appearance acceptance."""
from types import SimpleNamespace
from unittest.mock import Mock,patch
from threading import Event
from queue import Queue
from progress_studio.presentation.gui.boq_payment_dialog import BOQPaymentDialog


def test_revision_ack_blocks_destination_dialog():
    obj=SimpleNamespace(busy=False,review=SimpleNamespace(revision_required=True),ack=Mock(get=lambda:False))
    with patch('progress_studio.presentation.gui.boq_payment_dialog.filedialog.asksaveasfilename') as ask:
        BOQPaymentDialog.export(obj);ask.assert_not_called()


def test_busy_close_requests_cancel_without_destroy():
    obj=SimpleNamespace(busy=True,cancel=Event(),status=Mock(),destroy=Mock())
    BOQPaymentDialog.close(obj);assert obj.cancel.is_set();obj.destroy.assert_not_called()


def test_worker_only_receives_plain_review_payload():
    service=Mock();service.review.return_value='reviewed';events=Queue()
    obj=SimpleNamespace(busy=False,cancel=Event(),sync=Mock(),progress=Mock(),status=Mock(),service=service,events=events,after=Mock(),poll=Mock())
    BOQPaymentDialog.start(obj,'review',('session','payment',''))
    action,value,error=events.get(timeout=5)
    assert (action,value,error)==('review','reviewed',None)
    assert service.review.call_args.args==('session','payment','')
    assert obj.busy and obj.after.called
