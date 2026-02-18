.PHONY: setup dev-backend dev-frontend dev build docker-up docker-down

setup:
	cd backend && pip install -r requirements.txt
	cd frontend && npm install

dev-backend:
	cd backend && uvicorn main:app --reload --host 0.0.0.0 --port 8000

dev-frontend:
	cd frontend && npm run dev

dev:
	$(MAKE) dev-backend & $(MAKE) dev-frontend & wait

build:
	cd frontend && npm run build

docker-up:
	docker compose up --build -d

docker-down:
	docker compose down
