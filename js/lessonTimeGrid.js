/**
 * Модуль школьной сетки времени уроков.
 *
 * Содержит две эталонные сетки:
 * - mondaySchedule: расписание звонков для понедельника
 * - tueToSatSchedule: расписание звонков для вторника–субботы
 *
 * Экспортирует функцию getLessonTime(dayOfWeek, lessonNumber, shift)
 */

const LessonTimeGrid = (() => {
    'use strict';

    // Сетка для понедельника
    const mondaySchedule = [
        { shift: 1, lessonNumber: 1, startTime: '8:35',  endTime: '9:15'  },
        { shift: 1, lessonNumber: 2, startTime: '9:30',  endTime: '10:10' },
        { shift: 1, lessonNumber: 3, startTime: '10:25', endTime: '11:05' },
        { shift: 1, lessonNumber: 4, startTime: '11:25', endTime: '12:05' },
        { shift: 1, lessonNumber: 5, startTime: '12:25', endTime: '13:05' },
        { shift: 1, lessonNumber: 6, startTime: '13:20', endTime: '14:00' },
        { shift: 1, lessonNumber: 7, startTime: '14:10', endTime: '14:50' },
        { shift: 1, lessonNumber: 8, startTime: '14:55', endTime: '15:35' },
    ];

    // Сетка для вторника–субботы
    const tueToSatSchedule = [
        { shift: 1, lessonNumber: 1, startTime: '8:00',  endTime: '8:40'  },
        { shift: 1, lessonNumber: 2, startTime: '8:55',  endTime: '9:35'  },
        { shift: 1, lessonNumber: 3, startTime: '9:50',  endTime: '10:30' },
        { shift: 1, lessonNumber: 4, startTime: '10:50', endTime: '11:30' },
        { shift: 1, lessonNumber: 5, startTime: '11:50', endTime: '12:30' },
        { shift: 1, lessonNumber: 6, startTime: '12:45', endTime: '13:25' },
        { shift: 1, lessonNumber: 7, startTime: '13:35', endTime: '14:15' },
        { shift: 1, lessonNumber: 8, startTime: '14:20', endTime: '15:00' },
    ];

    // Маппинг коротких названий дней недели
    const DAY_ABBREVS = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];
    const MONDAY_DAYS = ['Пн'];
    const TUE_SAT_DAYS = ['Вт', 'Ср', 'Чт', 'Пт', 'Сб'];

    /**
     * Получить время урока по дню недели и номеру урока.
     *
     * @param {string} dayOfWeek - День недели (Пн, Вт, Ср, Чт, Пт, Сб, Вс)
     * @param {number} lessonNumber - Номер урока (1-8)
     * @param {number} [shift=1] - Номер смены (по умолчанию 1)
     * @returns {{startTime: string, endTime: string, shift: number, timeSource: string, found: boolean}}
     */
    function getLessonTime(dayOfWeek, lessonNumber, shift) {
        shift = shift || 1;

        if (!dayOfWeek || !lessonNumber) {
            return {
                startTime: null,
                endTime: null,
                shift: shift,
                timeSource: 'не определено',
                found: false
            };
        }

        const day = dayOfWeek.trim();
        let schedule;

        if (MONDAY_DAYS.includes(day)) {
            schedule = mondaySchedule;
        } else if (TUE_SAT_DAYS.includes(day)) {
            schedule = tueToSatSchedule;
        } else {
            // Воскресенье или неизвестный день
            return {
                startTime: null,
                endTime: null,
                shift: shift,
                timeSource: 'не определено',
                found: false
            };
        }

        const lesson = schedule.find(
            l => l.lessonNumber === lessonNumber && l.shift === shift
        );

        if (!lesson) {
            return {
                startTime: null,
                endTime: null,
                shift: shift,
                timeSource: 'не определено',
                found: false
            };
        }

        return {
            startTime: lesson.startTime,
            endTime: lesson.endTime,
            shift: lesson.shift,
            timeSource: 'школьная сетка',
            found: true
        };
    }

    /**
     * Получить полную сетку для отображения в UI.
     */
    function getMondaySchedule() {
        return mondaySchedule.slice();
    }

    function getTueToSatSchedule() {
        return tueToSatSchedule.slice();
    }

    function getDayAbbrevs() {
        return DAY_ABBREVS.slice();
    }

    return {
        getLessonTime,
        getMondaySchedule,
        getTueToSatSchedule,
        getDayAbbrevs,
        MONDAY_DAYS,
        TUE_SAT_DAYS,
        DAY_ABBREVS
    };
})();
