# -*- coding: utf-8 -*-
import http.client
import json
import base64


def auth(login_user: bytes, password_user: bytes):
    """
    Аутентификация

    :param login_user: логин пользователя
    :param password_user: пароль пользователя
    :return: dict
    """
    headers = {"Authorization": b"Basic " + base64.b64encode(login_user + b":" + password_user)}
    connection = http.client.HTTPSConnection("cn-ck11-web-ep.oducn.so:9443")
    connection.request("POST",
                       "/auth/app/token",
                       "",
                       headers)

    response = connection.getresponse()
    print(f"Status: {response.status} and reason: {response.reason}")
    responceBody = response.read()
    responseData = json.loads(responceBody)
    headers = {
        "Authorization": f"{responseData['token_type']} {responseData['access_token']}",
        "Content-Type": "application/json"
    }
    return headers, connection
