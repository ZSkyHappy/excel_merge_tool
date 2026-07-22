from __future__ import annotations

import argparse
import ctypes
import os
import threading
import time
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from merge_engine import MergeConfig, merge_workbooks


class ProcessMemoryCounters(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_ulong),
        ("PageFaultCount", ctypes.c_ulong),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
    ]


def current_working_set() -> int:
    if os.name != "nt":
        return 0
    counters = ProcessMemoryCounters()
    counters.cb = ctypes.sizeof(counters)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    kernel32.GetCurrentProcess.restype = ctypes.c_void_p
    psapi.GetProcessMemoryInfo.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(ProcessMemoryCounters),
        ctypes.c_ulong,
    ]
    psapi.GetProcessMemoryInfo.restype = ctypes.c_int
    ok = psapi.GetProcessMemoryInfo(
        kernel32.GetCurrentProcess(),
        ctypes.byref(counters),
        counters.cb,
    )
    return int(counters.WorkingSetSize) if ok else 0


def create_input(path: Path, start: int, count: int) -> None:
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("数据")
    headers = ["编号", "姓名", "主讲老师"] + [f"字段{index}" for index in range(4, 18)]
    sheet.append(headers)
    for identifier in range(start, start + count):
        sheet.append(
            [identifier, f"用户{identifier}", "性能老师"]
            + [identifier * index for index in range(4, 18)]
        )
    workbook.save(path)
    workbook.close()


def run_benchmark(total_rows: int, file_count: int) -> None:
    if total_rows < 1 or file_count < 1:
        raise ValueError("行数和文件数必须大于 0。")

    with TemporaryDirectory(prefix="excel_merge_benchmark_") as temp:
        root = Path(temp)
        input_dir = root / "input"
        output_dir = root / "output"
        input_dir.mkdir()
        base_count, remainder = divmod(total_rows, file_count)

        print(f"生成 {total_rows:,} 行 × 17 列测试数据……")
        cursor = 1
        generation_started = time.perf_counter()
        for index in range(file_count):
            count = base_count + (1 if index < remainder else 0)
            create_input(input_dir / f"测试_{index + 1:02d}.xlsx", cursor, count)
            cursor += count
        generation_elapsed = time.perf_counter() - generation_started

        stop_sampling = threading.Event()
        peak_working_set = current_working_set()

        def sample_memory() -> None:
            nonlocal peak_working_set
            while not stop_sampling.wait(0.05):
                peak_working_set = max(peak_working_set, current_working_set())

        sampler = threading.Thread(target=sample_memory, daemon=True)
        sampler.start()
        merge_started = time.perf_counter()
        summary = merge_workbooks(MergeConfig(input_dir, output_dir, "性能测试结果.xlsx"))
        merge_elapsed = time.perf_counter() - merge_started
        stop_sampling.set()
        sampler.join(timeout=1)

        counted_rows = 0
        for output in summary.output_files:
            workbook = load_workbook(output, read_only=True, data_only=True)
            sheet = workbook["合并结果"]
            counted_rows += sum(1 for _ in sheet.iter_rows(min_row=2, values_only=True))
            workbook.close()
        if counted_rows != total_rows:
            raise RuntimeError(f"验收失败：预期 {total_rows:,} 行，实际 {counted_rows:,} 行。")

        print(f"输入生成耗时：{generation_elapsed:.2f} 秒")
        print(f"合并耗时：{merge_elapsed:.2f} 秒")
        print(f"处理速度：{total_rows / max(merge_elapsed, 0.001):,.0f} 行/秒")
        if peak_working_set:
            print(f"峰值工作集：{peak_working_set / 1024 / 1024:.1f} MiB")
        else:
            print("峰值工作集：当前平台未采集")
        print(f"输出文件数：{len(summary.output_files)}")
        print("结果复核：通过")


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel 合并工具性能验收")
    parser.add_argument("--rows", type=int, default=100_000, help="总数据行数")
    parser.add_argument("--files", type=int, default=2, help="输入文件数")
    args = parser.parse_args()
    run_benchmark(args.rows, args.files)


if __name__ == "__main__":
    main()

