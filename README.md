# Generic Backend With Node Typescript Typeorm

## Local Setup

1. Add .env file with the following variables.

   ```
   JWT_SECRET=
   DATABASE_NAME=
   DATABASE_HOST=localhost
   DATABASE_PORT=3306
   DATABASE_USER_NAME=
   DATABASE_PASSWORD=
   NODE_ENV=development
   COOKIES_DOMAIN=
   AWS_REGION=ap-south-1
   AWS_ACCESS_KEY=
   AWS_SECRET_KEY=
   SUBDOMAIN_OFFSET=
   HOST=
   ```

2. go to `src/data-source.ts` and uncomment `synchronize: true,`. This will synchronize the entities with the database and tables will be created in the local database.
3. run `npm run typeorm migration:run` to add the basic data with a sample user.
4. run `npm run start` run the project.

## Migrations for Database

- After making changes in the entity run `npm run typeorm migration:generate -n ./src/migration/{FILE_NAME}` to generate the migration file automatically.
- After this run `npm run typeorm migration:run` to apply the changes to the database.

## Git Branching

- `main` - stable production code.
- `dev` - latest dev code.
- `feature/feature-name` - feature related code. New feature branches will be based on `dev` branch. After feature is ready it will be merged in the `dev`.
