import { EntityManager } from "typeorm";
import {
  SearchUserSchema,
  UpdateUserSchema,
  UserSchema,
} from "../schemas/user";
import { ResourceNotFoundError, ValidationFailedError } from "../common/errors";
import { User } from "../entity/User";
import { AppDataSource } from "../data-source";
import { defaultPageNumber, defaultPageSize } from "../common/constants";

class UserService {
  getAllUsers = async (query: SearchUserSchema) => {
    const userRepository = await AppDataSource.getRepository(User);
    const pageSize = query.size || defaultPageSize;
    const pageNumber = query.page || defaultPageNumber;

    let qb = userRepository
      .createQueryBuilder("user")
      .select(["user.id", "user.firstName", "user.lastName", "user.email"]);

    if (query.order) {
      qb = qb.orderBy("user.id", query.order);
    }

    if (query.page) {
      qb.skip((+query.page - 1) * +pageSize);
    } else {
      qb.skip(+pageNumber);
    }

    qb.take(+pageSize);

    const users = qb.getMany();

    return users;
  };

  getUserById = async (id: number) => {
    const userRepository = await AppDataSource.getRepository(User);

    const whereClause = {
      id: id,
    };

    const user = await userRepository
      .createQueryBuilder("user")
      .where(whereClause)
      .select(["user.id", "user.firstName", "user.lastName", "user.email"])
      .getOne();

    return user;
  };

  getUserProfile = async (id: number) => {
    const userRepository = await AppDataSource.getRepository(User);

    const whereClause = {
      id: id,
    };

    const user = await userRepository
      .createQueryBuilder("user")
      .where(whereClause)
      .leftJoin("user.organisation", "org")
      .select([
        "user.id",
        "user.name",
        "user.email",
        "user.isVerified",
        "user.role",
        "user.phone",
        "org.name",
      ])
      .getOne();

    return user;
  };

  createUser = async (
    transactionEntityManager: EntityManager,
    user: UserSchema
  ) => {
    const userRepository = await transactionEntityManager.getRepository(User);

    const duplicateEmailUser = await userRepository.findOneBy({
      email: user.email,
    });

    if (duplicateEmailUser) {
      throw new ValidationFailedError(`user already exists`);
    }

    const dbUser = new User();
    dbUser.email = user.email;
    dbUser.firstName = user.firstName;
    dbUser.lastName = user.lastName;

    await transactionEntityManager.save(dbUser);

    return;
  };

  updateUser = async (
    transactionEntityManager: EntityManager,
    user: UpdateUserSchema,
    id: number
  ) => {
    const userRepository = await transactionEntityManager.getRepository(User);

    const currentUser = await userRepository.findOneBy({
      id: id,
    });

    if (currentUser) {
      try {
        await transactionEntityManager.update(User, { id: id }, user);
        return;
      } catch (error) {
        throw new Error(error.message);
      }
    }

    throw new ResourceNotFoundError(`user doesn't exists.`);
  };

  deleteUser = async (transactionEntityManager: EntityManager, id: number) => {
    const userRepository = await transactionEntityManager.getRepository(User);

    const currentUser = await userRepository.findOneBy({
      id: id,
    });

    if (currentUser) {
      await transactionEntityManager.softDelete(User, { id: id });
      return;
    }

    throw new Error("user doesn't exists.");
  };
}

export default UserService;
