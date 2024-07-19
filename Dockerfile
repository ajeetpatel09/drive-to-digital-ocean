# syntax=docker/dockerfile:1

ARG NODE_VERSION=18.16.0

# Base image
FROM node:${NODE_VERSION}-alpine as base
WORKDIR /usr/src/app
COPY package*.json ./

# Install production dependencies
FROM base as deps
COPY package*.json yarn.lock ./
RUN yarn install --production --frozen-lockfile

# Build application
FROM deps as build
COPY . .
RUN yarn install --frozen-lockfile && yarn run build

# Final production image
FROM base as final
ENV NODE_ENV production
USER node

# Copy dependencies and compiled JS from previous stages
COPY --from=deps /usr/src/app/node_modules ./node_modules
COPY --from=build /usr/src/app/build .

EXPOSE 8000
CMD node src/index.js
