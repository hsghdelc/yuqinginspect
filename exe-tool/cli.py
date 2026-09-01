import argparse

from processor import process_file


def main():
    parser = argparse.ArgumentParser(description="舆情质检辅助工具命令行入口")
    parser.add_argument("input", help="源 Excel 文件路径")
    parser.add_argument("-o", "--output-dir", default=None, help="输出目录")
    parser.add_argument("-i", "--inspector", default="未命名质检员", help="质检人员")
    args = parser.parse_args()

    result = process_file(args.input, args.output_dir, args.inspector)
    print("处理完成")
    print("专项质检:", result["special_name"])
    print("专项命中:", result["special_count"])
    print("舆情提醒复核:", result["reminder_count"])
    print("无效复核:", result["invalid_count"])
    print("超时:", result["overtime_count"])
    print("日报送文本:", result["report_text"])
    print("输出文件:", result["output_path"])


if __name__ == "__main__":
    main()
