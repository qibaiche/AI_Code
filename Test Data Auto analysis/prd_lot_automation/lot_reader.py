import logging
from pathlib import Path
from typing import List


LOGGER = logging.getLogger(__name__)


def read_lots(lots_path: Path) -> List[str]:
    """读取 LOT 列表，支持 # 注释与去重。"""
    unique_lots = []
    seen = set()
    with open(lots_path, "r", encoding="utf-8") as fp:
        for line in fp:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped not in seen:
                unique_lots.append(stripped)
                seen.add(stripped)

    LOGGER.info("LOT 列表读取完成：%s 个（原文件：%s）", len(unique_lots), lots_path)
    if not unique_lots:
        raise ValueError(f"LOT 列表为空，请检查 {lots_path}")
    return unique_lots


def split_batches(lots: List[str], batch_size: int | None) -> List[List[str]]:
    if batch_size is None or batch_size <= 0:
        return [lots]
    batches: List[List[str]] = []
    for i in range(0, len(lots), batch_size):
        batches.append(lots[i : i + batch_size])
    LOGGER.debug("LOT 列表分批完毕：%s 批，每批最多 %s", len(batches), batch_size)
    return batches

