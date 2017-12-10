"""Program to read in popular reddit.com posts, determine 
the Pareto principle's applicability, graph results in Excel"""
# Based on this video: https://youtu.be/fCn8zs912OE?t=4m20s
# More on Pareto (80/20) principle here: https://en.wikipedia.org/wiki/Pareto_principle

import praw  # Python Reddit API Wrapper
import xlsxwriter  # Python to Excel module



# Loop to iterate over all non-negative top_level comments, increment upvote counter, record non-negative comments
# Also writes all non-negative values to the xlsx worksheet
# format.set_font_size(font_size=14)


def read_comments(subreddit):
    # Xlsx initialization, with formatting
    workbook = xlsxwriter.Workbook('8020.xlsx')
    submissions = reddit.subreddit(subreddit).top('all', limit=3)
    sheetcount = 1
    for submission in submissions:
        name = "TopPost" + str(sheetcount)
        sheetcount += 1
        worksheet = workbook.add_worksheet(name)
        format = workbook.add_format({'bold': True})
        format.set_font_size(font_size=18)
        format.set_font_name('Segoe UI')
        worksheet.write('A1', '8020 Values', format)
        row = 1
        col = 0
        order = 1
        print(submission.title)
        upvotetotal = 0
        count = 0
        # Replaces all MoreComments objects in the CommentForest; essentially loads all comments first
        submission.comments.replace_more(limit=0)
        for top_comment in submission.comments:
            if top_comment.ups > 0:
                worksheet.write(row, col, order, format)
                worksheet.write(row, col + 1, top_comment.ups, format)
                upvotetotal += top_comment.ups
                count += 1
                order += 1
                row += 1
        top = count * .20
        upvote = 0
        loop = 0
        for top_comment in submission.comments:
            loop += 1
            if loop < top:
                upvote += top_comment.ups
            else:
                break
        worksheet.write(row, col, 'Upvotes: ', format)
        worksheet.write(row, col + 1, upvotetotal, format)
        row += 1
        worksheet.write(row, col, 'Comments: ', format)
        worksheet.write(row, col + 1, count, format)

        column = workbook.add_chart({'type': 'column'})
        column.add_series({
            'categories': '=' + name + '!$A$2:$A$' + str(count),
            'values': '=' + name + '!$B$2:$B$' + str(count)
        })
        column.set_title({'name': 'Comment Rank vs. Upvotes'})

        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({
            'categories': '=' + name + '!$A$2:$A$' + str(count),
            'values': '=' + name + '!$B$2:$B$' + str(count)
        })
        chart.set_x_axis({'name': 'Comment Rank'})
        chart.set_y_axis({'name': 'Upvotes'})
        chart.set_style(10)

        column.combine(chart)
        worksheet.insert_chart('D2', column, {'x_offset': 25, 'y_offset': 10, 'x_scale': 3, 'y_scale': 2})

    # Close the Xlsx
    workbook.close()
    return


read_comments("nba")
