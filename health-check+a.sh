#!/bin/bash
# Usage: ./health-check.sh
# 引数なし

TODAY="`date +%m%d`"
MAILDIR="/mnt/z/path/to/health-check/"
# MAILDIR="./"
EMAIL_BODY="email-body.txt"
EMAIL_BODY_UTF8="email-body_utf8.txt"
EMAIL_ADDRERSS="email-address_2.txt"
# EMAIL_BODY_HTML = EMAIL_BODY_HTML1 + EMAIL_BODY_HTML2 + EMAIL_BODY_HTML3
EMAIL_BODY_HTML="email-body.html"
EMAIL_BODY1_HTML="emailbody1.html"
EMAIL_BODY2_HTML="emailbody2.html"
EMAIL_BODY3_HTML="emailbody3.html"
SUBJECT="【health-check】本日の入力状況"
FROM="xxx@xxx.com"
TO="yyy@xxx.com"
# TO=${FROM}
CC1=${FROM}
CC2="zzz@xxx.com"
# CC2=${FROM}

export LC_CTYPE=ja_JP.UTF-8

Send_html() {
	TO=$1
	CC1=$2
	CC2=$3
    SUBJECT=$4
    HTMLFILE=$5
    (echo "To: ${TO}"; echo "From: ${FROM}"; echo "Cc: ${CC1}"; echo "Cc: ${CC2}"; echo "Subject: ${SUBJECT}"; echo 'Mime-Version: 1.0'; echo 'Content-Type: text/html'; echo ) \
    | cat - ${HTMLFILE} | /usr/sbin/sendmail -t
}

# health-check.xlsx 処理
python3 /home/vuls/health-check/health-check.py

# ファイル処理
ls ${MAILDIR}${EMAIL_BODY} >/dev/null 2>&1
if [ $? -ne 0 ]
then
	echo "No such a text file: "${MAILDIR}${EMAIL_BODY}
	echo ""
else
	echo "Text file exists: " ${MAILDIR}${EMAIL_BODY}
	# Shift_JISからUTF-8 に変換
	# nkf -w --overwrite ${MAILDIR}${EMAIL_BODY}
	nkf -w ${MAILDIR}${EMAIL_BODY} > ${MAILDIR}${EMAIL_BODY_UTF8}

	# HTML作成
	# 空行を削除,      行頭に'<tr>\n<td>',    行末に'</td>\n</tr>',     1つ目の'-'を'<br>',    '\t'を'</td>\n<td>', 背景色変更
	sed -e '/^$/d' -e 's/^/<tr>\n<td>/g' -e 's/$/<\/td>\n<\/tr>/g' -e 's/-/<br>/' -e 's/\t/<\/td>\n<td>/g' -e 's/<td>▲未入力/<td bgcolor="#ffffc0">▲未入力/g' ${MAILDIR}${EMAIL_BODY_UTF8} > ${MAILDIR}${EMAIL_BODY2_HTML}
    cat ${MAILDIR}${EMAIL_BODY2_HTML}

	# ファイル連結
	cat ${MAILDIR}${EMAIL_BODY1_HTML} ${MAILDIR}${EMAIL_BODY2_HTML} ${MAILDIR}${EMAIL_BODY3_HTML} > ${MAILDIR}${EMAIL_BODY_HTML}
    cat ${MAILDIR}${EMAIL_BODY_HTML}

    # メール送信
	Send_html ${TO} ${CC1} ${CC2} ${SUBJECT} ${MAILDIR}${EMAIL_BODY_HTML}
fi

# 移動
mv -f ${MAILDIR}email-body*.txt ${MAILDIR}done

echo "Shell done!"
