# -*- coding: utf-8 -*-
"""Mebel360° qo‘shimcha modullarini Gunicorn workeriga ulash."""


def post_worker_init(worker):
    """Flask ilovasi yuklangach, birinchi so‘rovdan oldin kengaytmani ulaydi."""
    try:
        import app as app_module
        from mebel360_extension import patch_app

        patch_app(app_module)
        worker.log.info("Mebel360 buyurtma o‘chirish va avans siyosati moduli ulandi")
    except Exception:
        worker.log.exception("Mebel360 qo‘shimcha modulini ulashda xato")
        raise
