"""
导出第六章 表6.8、6.9、6.10 为三线表图片（PPT 展示用）
Python 环境：Anaconda
依赖：matplotlib（Anaconda 自带）
"""
import matplotlib.pyplot as plt
import matplotlib
import os

# 设置中文字体（Windows 系统常见中文字体）
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'KaiTi']
matplotlib.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'imgs', 'tables')
os.makedirs(OUTPUT_DIR, exist_ok=True)


def draw_three_line_table(title, headers, rows, filename, col_widths=None):
    """
    绘制三线表并保存为 PNG 图片。

    Parameters
    ----------
    title : str       - 表标题
    headers : list    - 列名列表
    rows : list[list] - 数据行
    filename : str    - 输出文件名
    col_widths : list - 各列相对宽度（可选）
    """
    n_rows = len(rows)
    n_cols = len(headers)

    if col_widths is None:
        col_widths = [1.0] * n_cols
    total_w = sum(col_widths)
    col_widths = [w / total_w for w in col_widths]

    # 计算图片尺寸
    cell_height = 0.55
    header_height = 0.65
    fig_height = (n_rows * cell_height + header_height + 1.8) * 0.35
    fig_width = max(total_w * 0.9, 7.0)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis('off')

    # 标题
    ax.set_title(title, fontsize=14, fontweight='bold', pad=18)

    # 表格区域
    table_top = 0.92
    table_bottom = 0.08
    table_left = 0.05
    table_right = 0.95

    col_x = [table_left]
    for w in col_widths:
        col_x.append(col_x[-1] + w * (table_right - table_left))
    col_x_mid = [(col_x[i] + col_x[i + 1]) / 2 for i in range(n_cols)]

    row_height = (table_top - table_bottom) / (n_rows + 1.5)
    header_y = table_top
    data_ys = [table_top - row_height * (i + 1.2) for i in range(n_rows)]

    # ---- 三线 ----
    lw = 1.5
    # 顶线
    ax.plot([table_left, table_right], [header_y + row_height * 0.4] * 2,
            linewidth=lw, color='black', transform=ax.transAxes)
    # 栏目线（header 下方）
    ax.plot([table_left, table_right], [header_y - row_height * 0.35] * 2,
            linewidth=0.8, color='black', transform=ax.transAxes)
    # 底线
    ax.plot([table_left, table_right], [data_ys[-1] - row_height * 0.4] * 2,
            linewidth=lw, color='black', transform=ax.transAxes)

    # ---- 表头 ----
    for j, h in enumerate(headers):
        ax.text(col_x_mid[j], header_y, h,
                ha='center', va='center', fontsize=11, fontweight='bold',
                transform=ax.transAxes)

    # ---- 数据行 ----
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            ax.text(col_x_mid[j], data_ys[i], str(cell),
                    ha='center', va='center', fontsize=10,
                    transform=ax.transAxes)

    plt.tight_layout()
    out_path = os.path.join(OUTPUT_DIR, filename)
    fig.savefig(out_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"Done: {out_path}")


# =====================================================================
# 表 6.8  安全性测试用例表
# =====================================================================
draw_three_line_table(
    title='表 6.8  安全性测试用例表',
    headers=['编号', '测试用例', '测试步骤', '预期结果', '测试结果'],
    rows=[
        ['TC24', '重大问题可见性',
         '分别用管理员、项目成员\n账号登录',
         '管理员看到所有重大问题，\n项目成员看到和自己相关\n的重大问题',
         '符合预期'],
        ['TC25', '重大问题新增权限',
         '使用管理员账号和普通\n用户账号新增问题',
         '管理员可以新增成功，\n普通用户提示无权限',
         '符合预期'],
    ],
    filename='table_6_8_safety_testcase.png',
    col_widths=[0.6, 1.4, 2.0, 2.2, 1.0],
)

# =====================================================================
# 表 6.9  普通分页响应时间
# =====================================================================
draw_three_line_table(
    title='表 6.9  普通分页响应时间（单位：ms）',
    headers=['页号', '第1次', '第2次', '第3次', '第4次', '第5次', '平均值'],
    rows=[
        ['1',    '18',    '10',    '5',     '16',    '12',    '12.2'],
        ['1K',   '8453',  '5831',  '6626',  '6500',  '6532',  '6788.4'],
        ['2K',   '13874', '12099', '14021', '12327', '13023', '13068.8'],
        ['3K',   '22061', '20695', '18858', '21390', '21262', '20853.2'],
        ['4K',   '29469', '29996', '29816', '28479', '28327', '29217.4'],
        ['5K',   '35338', '34108', '35862', '36300', '37072', '35736.0'],
    ],
    filename='table_6_9_pagination_before.png',
    col_widths=[0.6, 0.9, 0.9, 0.9, 0.9, 0.9, 1.0],
)

# =====================================================================
# 表 6.10  深分页优化后响应时间
# =====================================================================
draw_three_line_table(
    title='表 6.10  深分页优化后响应时间（单位：ms）',
    headers=['页号', '第1次', '第2次', '第3次', '第4次', '第5次', '平均值'],
    rows=[
        ['1',    '19',    '15',    '10',    '16',    '12',    '14.4'],
        ['1K',   '20',    '50',    '18',    '16',    '28',    '26.4'],
        ['2K',   '51',    '37',    '40',    '25',    '37',    '38.0'],
        ['3K',   '77',    '31',    '46',    '49',    '22',    '45.0'],
        ['4K',   '50',    '77',    '61',    '33',    '92',    '62.6'],
        ['5K',   '26',    '84',    '108',   '59',    '89',    '73.2'],
    ],
    filename='table_6_10_pagination_after.png',
    col_widths=[0.6, 0.9, 0.9, 0.9, 0.9, 0.9, 1.0],
)

# =====================================================================
# 表 6.11  查询执行计划对比
# =====================================================================
draw_three_line_table(
    title='表 6.11  查询执行计划对比',
    headers=['版本', 'type', 'key', 'Extra'],
    rows=[
        ['优化前', 'ALL',   'NULL',              'Using where; Using temporary;\nUsing filesort'],
        ['优化后', 'range', 'idx_wh_work_date',  'Using index condition;\nUsing temporary; Using filesort'],
    ],
    filename='table_6_11_sql_exec_plan.png',
    col_widths=[0.8, 0.8, 2.0, 3.0],
)


# =====================================================================
# 汇总表：深分页优化前后平均响应时间对比（PPT 重点展示）
# =====================================================================
draw_three_line_table(
    title='深分页优化前后平均响应时间对比（单位：ms）',
    headers=['页号', '优化前', '优化后', '提升倍数'],
    rows=[
        ['1',   '12.2',     '14.4',     '—'],
        ['1K',  '6788.4',   '26.4',     '257×'],
        ['2K',  '13068.8',  '38.0',     '344×'],
        ['3K',  '20853.2',  '45.0',     '463×'],
        ['4K',  '29217.4',  '62.6',     '467×'],
        ['5K',  '35736.0',  '73.2',     '488×'],
    ],
    filename='table_deep_page_summary.png',
    col_widths=[0.8, 1.2, 1.2, 1.2],
)

print("\nAll done! Images saved to:", os.path.abspath(OUTPUT_DIR))
