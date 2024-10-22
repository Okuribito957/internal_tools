import git
import xlwings

# Git本地仓库路径
repo_path = 'D:/BTP/qas/100_SAPBTP'
# 读取Excel绝对路径   C:/Users/Administrator/Desktop/temp/2/修正資産管理台帳.xlsm
excel_read_path = 'C:/Users/Administrator/Desktop/temp/2/修正資産管理台帳.xlsm'
# 保存Excel绝对路径
excel_save_path = 'C:/Users/Administrator/Desktop/temp/2/修正資産管理台帳_new2.xlsm'
# 本次变更的JIRA号或变更信息列名: ARO
draw_column = 'ARO'
# SHA
commit_sha = '9ed46ce85d9cde396f2cbe342bf0ed4b5417d804'


def get_commit_files(repo_path_inner, commit_sha_inner):
    """
    获取Git提交受影响的文件
    :param repo_path_inner: Git本地仓库路径
    :param commit_sha_inner: 提交的CommitSHA
    :return: 该次提交的变更文件列表
    """
    repo = git.Repo(repo_path_inner)
    commit_info = repo.commit(commit_sha_inner)
    files_info = commit_info.stats.files
    # 该次提交的文件列表
    files_changed_inner = []
    for file_path in files_info:
        files_changed_inner.append(file_path)
    return files_changed_inner


def get_excel_info(ws_inner):
    """
    获取Excel信息
    :param ws_inner: sheet页对象
    :return: 列信息字典
    """
    columns_dictionary = {}
    # first_row:管理番号
    column_values = ws_inner.range('C:C').value
    for i in range(len(column_values)):
        if column_values[i] is not None:
            columns_dictionary[column_values[i]] = i
    return columns_dictionary


def do_circle(wb_inner):
    """
    执行标记操作
    :return:
    """
    ws = wb_inner.sheets['資材一覧']
    # 单元格样式
    # font_name = wb_inner.sheets['プルダウンリスト'].range('G3').name
    columns_info = get_excel_info(ws)
    sum_success_files = 0
    sum_error_files = 0
    for file in files_changed:
        if columns_info.get(file) is not None:
            cell_name = draw_column + str(columns_info[file] + 1)
            ws.range(cell_name).value = "○"
            print("执行文件名：" + file + ",执行行数：" + str(columns_info[file]), end="\n")
            sum_success_files += 1
        else:
            print("不存在的文件名：" + file, end="\n")
            sum_error_files += 1
    print("共标记了" + str(sum_success_files) + "个文件,有" + str(sum_error_files) + "个文件不存在")
    # 保存文件
    wb_inner.save(excel_save_path)
    wb_inner.close()


if __name__ == '__main__':
    # 读取提交的文件
    files_changed = get_commit_files(repo_path, commit_sha)
    # 读取Excel
    wb = xlwings.Book(excel_read_path)
    # 标记操作
    if files_changed is not None:
        do_circle(wb)
    else:
        print("未获取到变更文件!", end="\n")
