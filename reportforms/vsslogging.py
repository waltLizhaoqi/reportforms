# -*- coding: utf-8 -*-
import logging.handlers
LOG_FILE = r'vspider.log'

handler = logging.handlers.RotatingFileHandler(LOG_FILE, maxBytes=1024 * 1024, backupCount=5, encoding='utf-8')  # 实例化handler
fmt = '%(asctime)s - %(levelname)s - %(message)s'

formatter = logging.Formatter(fmt)  # 实例化formatter
handler.setFormatter(formatter)  # 为handler添加formatter

logger = logging.getLogger('vspider')  # 获取名为tst的logger
logger.addHandler(handler)  # 为logger添加handler
logger.setLevel(logging.DEBUG)
