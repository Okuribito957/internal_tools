import warnings
import openpyxl
import git
# from deprecated.sphinx import deprecated

# Git本地仓库路径
repo_path = 'D:/BTP/qas/100_SAPBTP'
# Excel绝对路径
excel_path_read = 'C:/Users/Administrator/Desktop/temp/2/修正資産管理台帳.xlsm'
# 本次变更的JIRA号或变更信息列名: ARO
draw_column = 'ARO'
# SHA
commit_sha = '9ed46ce85d9cde396f2cbe342bf0ed4b5417d804'


def get_commit_files(repo_path, commit_sha):
    """
    获取Git提交受影响的文件
    :param repo_path: Git本地仓库路径
    :param commit_sha: 提交的CommitSHA
    :return: 该次提交的变更文件
    """
    repo = git.Repo(repo_path)
    commit_info = repo.commit(commit_sha)
    files_info = commit_info.stats.files
    # 该次提交的文件列表
    files_changed_inner = []
    for file_path in files_info:
        files_changed_inner.append(file_path)

    return files_changed_inner


# @deprecated(version='1.0', reason="openpyxl操作xlsm后会导致VBA损坏")
def get_excel_info(ws_inner):
    warnings.warn("openpyxl操作xlsm后会导致VBA损坏", DeprecationWarning)
    """
    获取Excel信息
    :param ws_inner: sheet页对象
    :return:
    """
    # first_row:管理番号
    first_row = ws_inner[1]
    row_dictionary = {}
    for row_cell in first_row:
        if row_cell.internal_value is not None:
            row_dictionary[row_cell.column_letter] = row_cell.internal_value
    # first_col: 文件名
    first_col = ws_inner["C"]
    col_dictionary = {}
    for col_cell in first_col:
        if col_cell.column is not None and type(col_cell) is not openpyxl.cell.cell.MergedCell:
            col_dictionary[col_cell.internal_value] = col_cell.row

    return row_dictionary, col_dictionary


if __name__ == '__main__':
    # 提交的文件
    files_changed = get_commit_files(repo_path, commit_sha)
    # row_info:行信息 col_info:列信息
    wb = openpyxl.load_workbook(excel_path_read)
    ws = wb["資材一覧"]
    row_info, col_info = get_excel_info(ws)
    for file in files_changed:
        if col_info.get(file) is not None:
            cell = draw_column + str(col_info.get(file))
            ws[cell].value = "○"
            print("执行文件名：" + file + ",执行行数：" + str(col_info.get(file)), end="\n")
    wb.save('C:/Users/Administrator/Desktop/temp/2/修正資産管理台帳5.xlsm')
