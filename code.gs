function doPost(e) {
    try {
        if (!e || !e.postData || !e.postData.contents) {
            return createJsonResponse_({
                ok: false,
                error: 'EMPTY_REQUEST',
                message: '전송된 데이터가 없습니다.'
            });
        }

        const data = JSON.parse(e.postData.contents);

        if (!data || typeof data !== 'object' || Array.isArray(data)) {
            return createJsonResponse_({
                ok: false,
                error: 'INVALID_REQUEST',
                message: '요청 형식이 올바르지 않습니다.'
            });
        }

        // 지원하는 요청만 처리
        if (
            data.action !== 'validateEntryCode' &&
            data.action !== 'submitAttendance'
        ) {
            return createJsonResponse_({
                ok: false,
                error: 'INVALID_ACTION',
                message: '지원하지 않는 요청입니다.'
            });
        }

        // 코드 확인 요청과 저장 요청 모두 서버에서 검사
        const validation = validateEntryCode_(data.entryCode);

        if (!validation.ok) {
            return createJsonResponse_(validation);
        }

        // 1차 요청: 코드만 확인하고 종료
        if (data.action === 'validateEntryCode') {
            return createJsonResponse_({
                ok: true,
                message: '입장 코드가 확인되었습니다.'
            });
        }

        // 2차 요청: 출석 내용 검사
        const studentInfo = String(data.studentInfo ?? '').trim();
        const resolution = String(data.resolution ?? '').trim();
        const bodyTemp = String(data.bodyTemp ?? '').trim();

        const moodTemp = Number(data.moodTemp);
        const emotions = Array.isArray(data.emotions)
            ? data.emotions
            : [];

        if (
            !studentInfo ||
            !resolution ||
            !bodyTemp ||
            data.moodTemp === undefined ||
            data.moodTemp === null ||
            String(data.moodTemp).trim() === '' ||
            !Number.isFinite(moodTemp) ||
            moodTemp < 0 ||
            moodTemp > 100 ||
            emotions.length < 1 ||
            emotions.length > 3 ||
            emotions.some(
                emotion =>
                    typeof emotion !== 'string' ||
                    !emotion.trim()
            )
        ) {
            return createJsonResponse_({
                ok: false,
                error: 'INVALID_ATTENDANCE',
                message: '출석 입력 내용을 확인해 주세요.'
            });
        }

        const spreadsheet = SpreadsheetApp.openById(
            ATTENDANCE_SPREADSHEET_ID
        );

        // 동시 제출 시 시트 생성·헤더 작성이 겹치지 않도록 처리
        const lock = LockService.getScriptLock();

if (!lock.tryLock(30000)) {
    return createJsonResponse_({
        ok: false,
        error: 'LOCK_TIMEOUT',
        message:
            '제출 요청이 몰려 저장 대기 시간을 초과했습니다. 잠시 후 다시 제출해 주세요.'
    });
}

try {
            let recordSheet = spreadsheet.getSheetByName('출석 기록');

            if (!recordSheet) {
                recordSheet = spreadsheet.insertSheet('출석 기록');
            }

            if (recordSheet.getLastRow() === 0) {
                recordSheet.appendRow([
                    '제출일시',
                    '학번/이름',
                    '기분온도',
                    '신체상태',
                    '감정단어',
                    '오늘의다짐'
                ]);

                recordSheet
                    .getRange(1, 1, 1, 6)
                    .setFontWeight('bold')
                    .setBackground('#dcfce7');
            }

            // 학생 입력이 시트 수식으로 해석되지 않도록 처리
            const asSheetText = value => {
                const text = String(value);

                return /^[=+\-@]/.test(text.trimStart())
                    ? "'" + text
                    : text;
            };

            recordSheet.appendRow([
                new Date(),
                asSheetText(studentInfo),
                moodTemp + '°C',
                asSheetText(bodyTemp),
                asSheetText(emotions.join(', ')),
                asSheetText(resolution)
            ]);

            SpreadsheetApp.flush();
        } finally {
            lock.releaseLock();
        }

        return createJsonResponse_({
            ok: true,
            message: '오늘의 마음 출석이 저장되었습니다.'
        });
    } catch (error) {
        console.error(error);

        return createJsonResponse_({
            ok: false,
            error: 'SERVER_ERROR',
            message:
                '요청을 처리하지 못했습니다. 저장 여부를 선생님께 확인해 주세요.'
        });
    }
}

function createJsonResponse_(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}
// 스프레드시트 주소에서 /d/ 다음부터 /edit 앞까지의 값
const ATTENDANCE_SPREADSHEET_ID = '1GEqDEnnRXtMY2YpoUFXv0dvwNi2HAi5DSOvMe13cASk';

function validateEntryCode_(enteredCode) {
    const spreadsheet = SpreadsheetApp.openById(
        ATTENDANCE_SPREADSHEET_ID
    );

    const codeSheet = spreadsheet.getSheetByName('입장 코드 입력');

    if (!codeSheet) {
        return {
            ok: false,
            error: 'CODE_NOT_CONFIGURED',
            message: '입장 코드가 설정되지 않았습니다. 선생님께 문의하세요.'
        };
    }

    const expectedCode = codeSheet
        .getRange('A1')
        .getDisplayValue()
        .trim();

    if (!expectedCode) {
        return {
            ok: false,
            error: 'CODE_NOT_CONFIGURED',
            message: '입장 코드가 설정되지 않았습니다. 선생님께 문의하세요.'
        };
    }

    const submittedCode = String(enteredCode ?? '').trim();

    if (!submittedCode || submittedCode !== expectedCode) {
        return {
            ok: false,
            error: 'INVALID_ENTRY_CODE',
            message: '입장 코드가 올바르지 않습니다. 선생님이 안내한 코드를 확인하세요.'
        };
    }

    return { ok: true };
}

function testEntryCodeValidation() {
    const results = {
        matching: validateEntryCode_('TEST-1234'),
        mismatching: validateEntryCode_('WRONG-CODE'),
        empty: validateEntryCode_('')
    };

    console.log(JSON.stringify(results));
}
