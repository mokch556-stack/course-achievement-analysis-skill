#!/usr/bin/env python3
"""
align_scores.py — 数据对齐脚本（余数分配法）
功能：当平时成绩报告单的子项和与录入系统平时分不一致时，对齐数据
以录入系统为准，用余数分配法重新分配子项分数

操作时间: 2026-06-28
操作agent: default
"""

import sys
import json
import argparse


def align_subs(subs, target_total):
    """
    余数分配法：将子项分数对齐到目标总分
    
    参数：
        subs: [xx, kt, jd, sc, ...] 原始子项分数列表
        target_total: 目标总分（录入系统的平时分）
    
    返回：
        [xx', kt', jd', sc', ...] 对齐后的整数子项分数
    """
    s = sum(subs)
    if abs(s - target_total) < 0.5:
        # 已经对齐，四舍五入取整即可
        return [int(round(v)) for v in subs]
    
    # 计算比例因子
    r = target_total / s
    
    # 按比例缩放
    scaled = [v * r for v in subs]
    
    # 取整数部分
    floors = [int(v) for v in scaled]
    
    # 计算余数
    left = int(round(target_total - sum(floors)))
    
    # 按小数部分从大到小排序
    decimals = [(i, v - int(v)) for i, v in enumerate(scaled)]
    decimals.sort(key=lambda x: x[1], reverse=True)
    
    # 将余数分配给小数部分最大的left项
    for i in range(min(left, len(decimals))):
        idx = decimals[i][0]
        floors[idx] += 1
    
    return floors


def verify_alignment(floors, target_total):
    """验证对齐结果"""
    actual = sum(floors)
    return actual == target_total, actual, target_total - actual


def main():
    parser = argparse.ArgumentParser(description='数据对齐脚本（余数分配法）')
    parser.add_argument('data_json', help='包含子项数据和目标总分的JSON文件')
    parser.add_argument('--output', '-o', help='输出对齐结果的JSON文件路径')
    parser.add_argument('--check', '-c', action='store_true', help='仅检查组一致性')
    args = parser.parse_args()
    
    with open(args.data_json, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    results = []
    errors = []
    
    for item in data:
        subs = item.get('subs', [])
        target = item.get('target', 0)
        name = item.get('name', '未知')
        
        # 检查
        s = sum(subs)
        if abs(s - target) < 0.5:
            results.append({
                "name": name,
                "status": "OK",
                "original": subs,
                "original_sum": s,
                "target": target,
                "diff": round(s - target, 2),
                "aligned": [int(round(v)) for v in subs],
                "note": "已自动对齐"
            })
            continue
        
        # 对齐
        aligned = align_subs(subs, target)
        ok, actual, diff = verify_alignment(aligned, target)
        
        r = {
            "name": name,
            "status": "OK" if ok else "FAIL",
            "original": subs,
            "original_sum": s,
            "target": target,
            "diff": round(s - target, 2),
            "aligned": aligned,
            "aligned_sum": actual,
            "note": f"余数分配法对齐" if ok else f"对齐失败（差{diff}）"
        }
        results.append(r)
        
        if not ok:
            errors.append(f"{name}: 对齐失败，差{diff}")
    
    output = {
        "total": len(data),
        "ok_count": sum(1 for r in results if r['status'] == 'OK'),
        "fail_count": sum(1 for r in results if r['status'] == 'FAIL'),
        "results": results,
        "errors": errors
    }
    
    if args.check:
        # 只输出检查结果
        print(f"\n📊 数据一致性检查")
        print(f"  · 检查总数: {output['total']}")
        print(f"  · 一致: {output['ok_count']}")
        print(f"  · 不一致: {output['fail_count']}")
        for r in results:
            if r['status'] == 'FAIL':
                print(f"    ❌ {r['name']}: 原始和{r['original_sum']} ≠ 目标{r['target']}")
        if output['fail_count'] == 0:
            print(f"  ✅ 全部一致")
        return
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2)
        print(f"✅ 对齐结果已保存: {args.output}")
    else:
        print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
