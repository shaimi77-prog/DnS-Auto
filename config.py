"""DnS Auto의 버전 정보와 실행 로그 설정."""
# Copyright (C) 2026 두부코드(DOOBOO_CODE)
# SPDX-License-Identifier: AGPL-3.0-only

import datetime
import logging
import os
import platform
import sys

VERSION = "v1.0.0"
PROGRAM_NAME = f"DnS Auto ({VERSION})"


def setup_detailed_logging():
    """실행 환경 정보를 기록하는 감사 로그 파일을 구성합니다."""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    log_fn = os.path.join(base_dir, f"DnS_Auto_AuditLog_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.txt")

    # 동일 프로세스에서 다시 초기화할 때 로그가 중복 기록되지 않도록 정리합니다.
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=log_fn,
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        encoding='utf-8'
    )

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    logging.getLogger('').addHandler(console)

    # 장애 분석에 필요한 최소 실행 환경을 기록합니다.
    logging.info("=" * 50)
    logging.info(f"프로그램: {PROGRAM_NAME}")
    logging.info(f"실행 일시: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info(f"실행 환경 OS: {platform.system()} {platform.release()} ({platform.architecture()[0]})")
    logging.info(f"파이썬 버전: {platform.python_version()}")
    logging.info(f"실행 모드: {'EXE 패키지 상태' if getattr(sys, 'frozen', False) else '스크립트 실행 상태'}")
    logging.info(f"실행 경로: {base_dir}")
    logging.info("=" * 50)

    return log_fn
