#!/usr/bin/env python3
"""
calc_matrix.py — 满分矩阵计算脚本
功能：根据教学大纲课点归属和考核构成，计算各目标各子项的满分分配
核心：支持 按课点比例分配（通识子项） 和 按项目归属分配（特定子项） 两种模式

操作时间: 2026-06-28
操作agent: default
"""

import sys
import json
import argparse
from copy import deepcopy


def calc_matrix(targets, assessment, course_points, subj_assignments, verbose=False):
    """
    计算满分矩阵
    
    参数：
        targets: [{"id": 1, "name": "目标1", "points": [1,2,3,4]}, ...]
            每个目标含id、name、points（所属课点列表）
        assessment: [{"name": "子项1", "full_score": 30, "type": "平时"}, ...]
            每个考核子项含name、full_score、type
        course_points: int，总课点数
        subj_assignments: {"子项1": "points", "子项2": "target:1", "子项3": "points", ...}
            "points" = 按课点比例分配
            "target:N" = 直接归属到目标N
            "manual" = 手动指定（后面可覆盖）
    
    返回：
        matrix: {"目标1": {"子项1": 13.44, ...}, ...}
        target_full_scores: {"目标1": 25, "目标2": 25, ...}
    """
    n_targets = len(targets)
    
    # 计算每个目标的课点数
    target_points = {f"目标{t['id']}": len(t.get('points', [])) for t in targets}
    total_points = sum(target_points.values())
    if total_points == 0:
        total_points = course_points
    
    # 各目标权重（按课点比例）
    target_weights = {k: v / total_points for k, v in target_points.items()}
    
    if verbose:
        print(f"\n📊 课点分析:")
        for t in targets:
            pts = t.get('points', [])
            print(f"  目标{t['id']}: {len(pts)}个课点 → 权重{target_weights[f'目标{t['id']}']:.2%}")
        print(f"  总课点: {total_points}")
    
    # 初始化矩阵
    matrix = {}
    for t in targets:
        matrix[f"目标{t['id']}"] = {a['name']: 0.0 for a in assessment}
    
    # 逐个考核子项分配
    for ass in assessment:
        name = ass['name']
        full = ass['full_score']
        assign = subj_assignments.get(name, 'points')
        
        if assign == 'points':
            # 按课点比例分配
            for t in targets:
                tid = f"目标{t['id']}"
                matrix[tid][name] = round(full * target_weights[tid], 2)
            if verbose:
                print(f"\n  📌 {name}({full}分): 按课点比例分配")
                for t in targets:
                    tid = f"目标{t['id']}"
                    print(f"     目标{t['id']}: {matrix[tid][name]:.2f}")
                    
        elif assign.startswith('target:'):
            # 直接归属到特定目标
            tid = int(assign.split(':')[1])
            tid_key = f"目标{tid}"
            matrix[tid_key][name] = full
            if verbose:
                print(f"\n  📌 {name}({full}分): 全部归属到目标{tid}")
                
        elif assign == 'manual':
            # 手动指定，保留0值待后续填写
            if verbose:
                print(f"\n  📌 {name}({full}分): 手动指定（待填写）")
    
    # 计算各目标满分
    target_full_scores = {}
    for t in targets:
        tid = f"目标{t['id']}"
        total = round(sum(matrix[tid].values()), 2)
        target_full_scores[tid] = total
    
    if verbose:
        print(f"\n📊 各目标满分:")
        for t in targets:
            tid = f"目标{t['id']}"
            print(f"  目标{t['id']}: {target_full_scores[tid]}分")
    
    # 验证：各目标满分之和应≈总满分
    total_full = sum(a['full_score'] for a in assessment)
    total_dist = sum(target_full_scores.values())
    if verbose:
        print(f"\n📊 验证:")
        print(f"  考核总满分: {total_full}")
        print(f"  各目标满分和: {total_dist:.2f}")
        if abs(total_dist - total_full) < 0.5:
            print(f"  ✅ 一致 (差{abs(total_dist - total_full):.2f})")
        else:
            print(f"  ⚠️ 不一致 (差{abs(total_dist - total_full):.2f}) — 需要调整")
    
    return matrix, target_full_scores


def adjust_matrix(matrix, target_full_scores, targets, adjustments=None):
    """
    调整矩阵使各目标满分与教学大纲规定一致
    
    adjustments: {"目标1": 25, "目标2": 25, ...}
        教学大纲规定的各目标满分值
    """
    if not adjustments:
        return matrix, target_full_scores
    
    adjusted_matrix = deepcopy(matrix)
    adjusted_scores = deepcopy(target_full_scores)
    
    # 找出哪些目标需要调整
    for tid, target_full in adjustments.items():
        current = adjusted_scores.get(tid, 0)
        diff = round(target_full - current, 2)
        
        if abs(diff) < 0.5:
            continue  # 无需调整
        
        # 找到该目标下按课点比例分配的"可调子项"（通常是笔记/课堂表现等）
        adjust_items = []
        for ass_name in adjusted_matrix[tid]:
            # 跳过期末项目等不可调项
            if '期末' in ass_name or '项目' in ass_name or '考试' in ass_name:
                continue
            adjust_items.append(ass_name)
        
        if adjust_items:
            # 按比例分配差额
            total_adj = sum(adjusted_matrix[tid][a] for a in adjust_items)
            if total_adj > 0:
                for a in adjust_items:
                    ratio = adjusted_matrix[tid][a] / total_adj
                    adjusted_matrix[tid][a] = round(adjusted_matrix[tid][a] + diff * ratio, 2)
        
        adjusted_scores[tid] = target_full
    
    return adjusted_matrix, adjusted_scores


def main():
    parser = argparse.ArgumentParser(description='满分矩阵计算脚本')
    parser.add_argument('config', help='课程配置JSON文件路径')
    parser.add_argument('--output', '-o', help='输出JSON文件路径（默认打印到控制台）')
    parser.add_argument('--verbose', '-v', action='store_true', help='显示详细分配过程')
    args = parser.parse_args()
    
    with open(args.config, 'r', encoding='utf-8') as f:
        cfg = json.load(f)
    
    matrix, target_scores = calc_matrix(
        cfg['targets'],
        cfg['assessment'],
        cfg.get('course_points', len(sum([t.get('points', []) for t in cfg['targets']], []))),
        cfg.get('subj_assignments', {}),
        verbose=args.verbose
    )
    
    result = {
        "matrix": matrix,
        "target_full_scores": target_scores
    }
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"✅ 满分矩阵已保存: {args.output}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
