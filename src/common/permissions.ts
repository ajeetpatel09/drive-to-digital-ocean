export const READ_USER = "READ_USER";
export const UPDATE_USER = "READ_USER";
export const DELETE_USER = "DELETE_USER";
export const CREATE_USER = "CREATE_USER";
export const READ_USERS = "READ_USERS";

type UserPermission =
  | "READ_USER"
  | "UPDATE_USER"
  | "DELETE_USER"
  | "READ_USERS"
  | "CREATE_USER";

export const userPermissions = [
  READ_USER,
  UPDATE_USER,
  DELETE_USER,
  CREATE_USER,
  READ_USERS,
];

type AdminPermission = "READ_USERS";

export const adminPermissions = [READ_USERS];

export type Permission = UserPermission | AdminPermission;
