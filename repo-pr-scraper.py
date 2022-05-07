from github import Github

import argparse
import xlsxwriter
import os


my_parser = argparse.ArgumentParser(description='Github repo name')
# Add the arguments
my_parser.add_argument('repo',
                       metavar='repo',
                       type=str,
                       help='the github repo name in format of USER_NAME/REPO_NAME. e.g. Trusted-AI/AIF360')

# Execute the parse_args() method
args = my_parser.parse_args()

PR_XLSX_HEADER = ["PR Number", "PR Title", "PR Open Date", "PR Link", "PR Label"]
ISSUE_XLSX_HEADER = ["Issue Number", "Issue Title", "Issue Open Date", "Issue Link", "Issue Label"]

workbook = xlsxwriter.Workbook(os.path.join(os.getcwd(), args.repo.split("/")[1]+'.xlsx'))
prs_worksheet = workbook.add_worksheet("prs")
issues_worksheet = workbook.add_worksheet("issues")

# using an access token
g = Github("chenhuiyeh", "wow!ti88316!")
repo = g.get_repo(args.repo)

pulls = []
for pr in repo.get_pulls(state='open', base='master'):
    pulls.append([pr.number, pr.title, pr.created_at, pr.html_url, pr.labels])

row = 0
col = 0
prs_worksheet.write_row(row,col,PR_XLSX_HEADER)
row += 1
for pr_num, pr_title, pr_open_date, pr_link, pr_label in pulls:
    prs_worksheet.write(row, col, pr_num)
    prs_worksheet.write(row, col + 1, pr_title)
    prs_worksheet.write(row, col + 2, pr_open_date.strftime("%Y/%m/%d"))
    prs_worksheet.write(row, col + 3, pr_link)
    pr_labels = [label.name for label in pr_label]
    prs_worksheet.write(row, col + 4, ",".join(pr_labels))
    row += 1
    col = 0

issues = []
for issue in repo.get_issues(state='open'):
    issues.append([issue.number, issue.title, issue.created_at, issue.html_url, issue.labels])
row = 0
col = 0
issues_worksheet.write_row(row,col,ISSUE_XLSX_HEADER)
row += 1
for issue_num, issue_title, issue_open_date, issue_link, issue_label in issues:
    issues_worksheet.write(row, col, issue_num)
    issues_worksheet.write(row, col + 1, issue_title)
    issues_worksheet.write(row, col + 2, issue_open_date.strftime("%Y/%m/%d"))
    issues_worksheet.write(row, col + 3, issue_link)
    issue_labels = [label.name for label in issue_label]
    issues_worksheet.write(row, col + 4, ",".join(issue_labels))
    row += 1
    col = 0
workbook.close()