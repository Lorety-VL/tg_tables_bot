# tg_tables_bot

## To start this bot

Clone this repo:

```sh
git clone git@github.com:Lorety-VL/tg_tables_bot.git
```

## Create .env file in root project dir with env TELEGRAM_BOT_TOKEN

example:

```sh
echo "TELEGRAM_BOT_TOKEN=your_telegram_bot_token" > .env
```

## Build docker

```sh
docker build -t tg_tables_bot .
```

## Run docker:

```sh
docker run -d --rm --name my-bot --env-file .env tg_tables_bot
```

## To stop this bot

```sh
docker kill tg_tables_bot
```