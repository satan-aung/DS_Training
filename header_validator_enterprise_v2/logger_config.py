import logging, os
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
 filename='logs/validation.log',
 level=logging.INFO,
 format='%(asctime)s %(levelname)s %(message)s'
)
