import { z } from "zod";
import { DefaultSearchParams } from "./comman";
// import { DefaultSearchParams } from "./common";

export const AddUser = z.object({
  firstName: z.string(),
  lastName: z.string(),
  email: z.string().email({ message: "Invalid email address" }),
});

export const SearchUser = DefaultSearchParams.extend({
  firstName: z.string().optional(),
  lastName: z.string().optional(),
});

export const UpdateUser = z.object({
  firstName: z.string(),
  lastName: z.string(),
});

export type UserSchema = z.infer<typeof AddUser>;
export type UpdateUserSchema = z.infer<typeof UpdateUser>;
export type SearchUserSchema = z.infer<typeof SearchUser>;
