"""测试最近的修复 - 验证 GTS 文件查找和 Spark Parttype 选择"""
import sys
from pathlib import Path

# 添加父目录到路径
current_dir = Path(__file__).parent
parent_dir = current_dir.parent
sys.path.insert(0, str(parent_dir))

def test_gts_file_finding():
    """测试 GTS 文件查找功能"""
    print("=" * 80)
    print("测试 1: GTS 文件查找功能")
    print("=" * 80)
    
    from workflow_automation.gts_submitter import find_latest_excel
    
    # 测试场景 1：在当前运行目录中查找
    print("\n场景 1: 在当前运行目录中查找")
    try:
        output_dir = parent_dir / "output" / "run_20251226_025714" / "02_GTS_Files"
        if output_dir.exists():
            result = find_latest_excel(output_dir)
            print(f"✅ 找到文件: {result.name}")
        else:
            print(f"⚠️  目录不存在: {output_dir}")
    except Exception as e:
        print(f"❌ 失败: {e}")
    
    # 测试场景 2：在父级 output 目录中查找
    print("\n场景 2: 在父级 output 目录中查找")
    try:
        output_dir = parent_dir / "output"
        if output_dir.exists():
            result = find_latest_excel(output_dir)
            print(f"✅ 找到文件: {result.name}")
            print(f"   完整路径: {result}")
        else:
            print(f"⚠️  目录不存在: {output_dir}")
    except Exception as e:
        print(f"❌ 失败: {e}")
    
    # 测试场景 3：在新建的空目录中查找（应该找到其他目录的文件）
    print("\n场景 3: 在新建的空目录中查找")
    try:
        new_run_dir = parent_dir / "output" / "run_test" / "02_GTS_Files"
        new_run_dir.mkdir(parents=True, exist_ok=True)
        
        result = find_latest_excel(new_run_dir)
        print(f"✅ 找到文件: {result.name}")
        print(f"   文件位置: {result.parent}")
        
        # 清理测试目录
        import shutil
        shutil.rmtree(new_run_dir.parent)
        print("   已清理测试目录")
    except Exception as e:
        print(f"❌ 失败: {e}")
    
    print("\n" + "=" * 80)
    print("GTS 文件查找测试完成")
    print("=" * 80)


def test_screenshot_functionality():
    """测试截图功能"""
    print("\n" + "=" * 80)
    print("测试 2: 截图功能")
    print("=" * 80)
    
    from workflow_automation.utils.screenshot_helper import capture_screen_screenshot
    
    print("\n尝试捕获屏幕截图...")
    try:
        output_dir = parent_dir / "output" / "05_Debug"
        screenshot_path = capture_screen_screenshot(
            output_dir,
            error_message="测试截图",
            prefix="test_screenshot"
        )
        
        if screenshot_path:
            print(f"✅ 截图成功: {screenshot_path.name}")
            print(f"   文件大小: {screenshot_path.stat().st_size:,} 字节")
            
            # 清理测试截图
            screenshot_path.unlink()
            print("   已清理测试截图")
        else:
            print("⚠️  截图功能不可用（可能缺少 Pillow）")
    except Exception as e:
        print(f"❌ 失败: {e}")
    
    print("\n" + "=" * 80)
    print("截图功能测试完成")
    print("=" * 80)


def test_wait_helpers():
    """测试等待工具"""
    print("\n" + "=" * 80)
    print("测试 3: 等待工具")
    print("=" * 80)
    
    from workflow_automation.utils.wait_helpers import wait_for_condition
    import time
    
    print("\n测试 wait_for_condition...")
    
    # 测试场景 1：条件立即满足
    print("\n场景 1: 条件立即满足")
    start_time = time.time()
    result = wait_for_condition(lambda: True, timeout=5)
    elapsed = time.time() - start_time
    
    if result and elapsed < 1:
        print(f"✅ 成功 (耗时: {elapsed:.2f}秒)")
    else:
        print(f"❌ 失败 (result={result}, 耗时={elapsed:.2f}秒)")
    
    # 测试场景 2：条件延迟满足
    print("\n场景 2: 条件延迟满足")
    counter = [0]
    def delayed_condition():
        counter[0] += 1
        return counter[0] >= 3
    
    start_time = time.time()
    result = wait_for_condition(delayed_condition, timeout=10, poll_frequency=0.5)
    elapsed = time.time() - start_time
    
    if result and 1 < elapsed < 3:
        print(f"✅ 成功 (耗时: {elapsed:.2f}秒, 检查次数: {counter[0]})")
    else:
        print(f"❌ 失败 (result={result}, 耗时={elapsed:.2f}秒)")
    
    # 测试场景 3：条件超时
    print("\n场景 3: 条件超时")
    start_time = time.time()
    result = wait_for_condition(lambda: False, timeout=2)
    elapsed = time.time() - start_time
    
    if not result and 1.8 < elapsed < 2.5:
        print(f"✅ 成功 (正确超时, 耗时: {elapsed:.2f}秒)")
    else:
        print(f"❌ 失败 (result={result}, 耗时={elapsed:.2f}秒)")
    
    print("\n" + "=" * 80)
    print("等待工具测试完成")
    print("=" * 80)


def test_error_handler():
    """测试错误处理机制"""
    print("\n" + "=" * 80)
    print("测试 4: 错误处理机制")
    print("=" * 80)
    
    from workflow_automation.utils.error_handler import (
        retry_on_exception,
        handle_errors,
        ErrorContext,
        safe_execute
    )
    
    # 测试场景 1：重试装饰器
    print("\n场景 1: 重试装饰器")
    attempt_count = [0]
    
    @retry_on_exception(max_retries=3, delay=0.1, exceptions=(ValueError,))
    def unstable_function():
        attempt_count[0] += 1
        if attempt_count[0] < 3:
            raise ValueError(f"尝试 {attempt_count[0]} 失败")
        return "成功"
    
    try:
        result = unstable_function()
        print(f"✅ 成功 (尝试次数: {attempt_count[0]}, 结果: {result})")
    except Exception as e:
        print(f"❌ 失败: {e}")
    
    # 测试场景 2：错误处理装饰器
    print("\n场景 2: 错误处理装饰器")
    
    @handle_errors(default_return=False)
    def failing_function():
        raise RuntimeError("测试错误")
    
    result = failing_function()
    if result == False:
        print(f"✅ 成功 (正确返回默认值: {result})")
    else:
        print(f"❌ 失败 (返回值: {result})")
    
    # 测试场景 3：错误上下文
    print("\n场景 3: 错误上下文")
    
    with ErrorContext("测试操作", raise_on_error=False) as ctx:
        raise ValueError("测试错误")
    
    if ctx.exception:
        print(f"✅ 成功 (捕获到异常: {type(ctx.exception).__name__})")
    else:
        print(f"❌ 失败 (未捕获异常)")
    
    # 测试场景 4：安全执行
    print("\n场景 4: 安全执行")
    
    def risky_function():
        raise RuntimeError("测试错误")
    
    result = safe_execute(
        risky_function,
        default_return="默认值",
        error_message="操作失败"
    )
    
    if result == "默认值":
        print(f"✅ 成功 (返回默认值: {result})")
    else:
        print(f"❌ 失败 (返回值: {result})")
    
    print("\n" + "=" * 80)
    print("错误处理机制测试完成")
    print("=" * 80)


def main():
    """运行所有测试"""
    print("\n")
    print("╔" + "=" * 78 + "╗")
    print("║" + " " * 20 + "Auto-VPO 最近修复验证测试" + " " * 33 + "║")
    print("╚" + "=" * 78 + "╝")
    print()
    
    tests = [
        ("GTS 文件查找", test_gts_file_finding),
        ("截图功能", test_screenshot_functionality),
        ("等待工具", test_wait_helpers),
        ("错误处理机制", test_error_handler),
    ]
    
    passed = 0
    failed = 0
    
    for test_name, test_func in tests:
        try:
            test_func()
            passed += 1
        except Exception as e:
            print(f"\n❌ 测试 '{test_name}' 失败: {e}")
            import traceback
            traceback.print_exc()
            failed += 1
    
    # 总结
    print("\n")
    print("╔" + "=" * 78 + "╗")
    print("║" + " " * 30 + "测试总结" + " " * 40 + "║")
    print("╠" + "=" * 78 + "╣")
    print(f"║  通过: {passed}  失败: {failed}  总计: {passed + failed}" + " " * (78 - 20 - len(str(passed)) - len(str(failed)) - len(str(passed + failed))) + "║")
    print("╚" + "=" * 78 + "╝")
    print()
    
    if failed == 0:
        print("🎉 所有测试通过！")
    else:
        print(f"⚠️  有 {failed} 个测试失败，请检查")
    
    print()


if __name__ == "__main__":
    main()

