# -*- coding: utf-8 -*-
from json import dumps


def post_json(uid: str, date_1: str, date_2: str, step_seconds: int = 1) -> str:
    """

    :param uid: UID - индентификационный номер
    :param date_1: 2022-01-25T10:05:00.000Z
    :param date_2: 2022-01-25T20:05:00.000Z
    :param step_seconds: шаг в секундах
    :return:
    """
    payload = dumps(
        {
            "uids": [uid],
            "fromTimeStamp": date_1,
            "toTimeStamp": date_2,
            "stepUnits": "seconds",
            "stepValue": step_seconds
        }
    )
    return payload


if __name__ == '__main__':
    print(post_json(uid='cc9dcfb8-f08b-49d7-b146-78a5b13bd4bc',
                    date_1='2022-01-25T10:05:00.000Z',
                    date_2='2022-01-25T10:05:00.000Z', step_seconds=1))
