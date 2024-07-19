import { userJwtTokenValidDays } from "../common/constants";
import { ResourceNotFoundError, ValidationFailedError } from "../common/errors";
import { AppDataSource } from "../data-source";
import { User } from "../entity/User";
import { LoginDataSchema, SignUpDataScehema } from "../schemas/auth";
import * as bcrypt from "bcrypt";
import * as jwt from "jsonwebtoken";
import UserService from "./users.service";
import { UserSchema } from "../schemas/user";
import { EntityManager } from "typeorm";

class AuthService {
  private _usersService = new UserService();

  login = async (loginData: LoginDataSchema) => {
    const userRepository = await AppDataSource.getRepository(User);
    const user = await userRepository
      .createQueryBuilder("user")
      .where({
        email: loginData.email,
      })
      .select([
        "user.id",
        "user.firstName",
        "user.lastName",
        "user.email",
        "user.password",
      ])
      .getOne();

    if (!user) {
      throw new ResourceNotFoundError(`user not found`);
    }

    const isMatching = user.password
      ? await bcrypt.compare(loginData.password, user.password)
      : false;

    if (!isMatching) {
      throw new ValidationFailedError(`wrong password`);
    }

    const jwtToken = jwt.sign(
      {
        id: user.id,
      },
      process.env.JWT_SECRET ?? "",
      { expiresIn: `${userJwtTokenValidDays}d` }
    );

    const newUserData = { ...user };
    delete newUserData.password;

    return { user: newUserData, jwtToken };
  };

  signup = async (
    transactionEntityManager: EntityManager,
    user: SignUpDataScehema
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
    dbUser.password = user.password;

    const savedUser = await transactionEntityManager.save(dbUser);

    const jwtToken = jwt.sign(
      {
        id: savedUser.id,
      },
      process.env.JWT_SECRET ?? "",
      { expiresIn: `${userJwtTokenValidDays}d` }
    );

    const newUserData = { ...savedUser };
    delete newUserData.password;

    return { user: newUserData, jwtToken };

    return { user: savedUser, jwtToken };
  };
}

export default AuthService;
