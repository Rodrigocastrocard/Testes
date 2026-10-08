import json
from dataclasses import dataclass
from functools import lru_cache

import boto3
from botocore.exceptions import ClientError


@dataclass(frozen=True)
class DBConfig:
    DBHost: str
    DBName: str
    DBUser: str
    DBPassword: str


@dataclass(frozen=True)
class SlackConfig:
    SlackBotToken: str


@dataclass(frozen=True)
class Config:
    DBConfig: DBConfig
    SlackConfig: SlackConfig


def get_secret_map(secret_name: str, region_name: str) -> dict:
    """
    Recupera um segredo do AWS Secrets Manager e retorna como dict.
    """
    try:
        client = boto3.client("secretsmanager", region_name=region_name)
        response = client.get_secret_value(
            SecretId=secret_name,
            VersionStage="AWSCURRENT",
        )
    except ClientError as e:
        raise RuntimeError(
            f"Erro ao buscar segredo '{secret_name}' na região '{region_name}'"
        ) from e

    secret_string = response.get("SecretString")
    if not secret_string:
        raise ValueError(
            f"Segredo '{secret_name}' não contém 'SecretString'"
        )

    try:
        return json.loads(secret_string)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"Segredo '{secret_name}' não contém JSON válido"
        ) from e


@lru_cache(maxsize=None)
def load_config(region_name: str) -> Config:
    """
    Carrega os segredos apenas uma vez por região durante a execução do processo.
    Se chamada novamente com a mesma região, retorna do cache sem nova ida ao AWS Secrets Manager.
    """
    db_secrets = get_secret_map("ams/db_credentials", region_name)
    slack_secrets = get_secret_map("ams/slack_credentials", region_name)

    db_config = DBConfig(
        DBHost=db_secrets["DB_HOST"],
        DBName=db_secrets["DB_NAME"],
        DBUser=db_secrets["DB_USER"],
        DBPassword=db_secrets["DB_PASSWORD"],
    )

    slack_config = SlackConfig(
        SlackBotToken=slack_secrets["SLACK_BOT_TOKEN"],
    )

    return Config(
        DBConfig=db_config,
        SlackConfig=slack_config,
    )