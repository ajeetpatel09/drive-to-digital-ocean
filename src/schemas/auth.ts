import { z } from "zod";

export const LoginData = z.object({
  email: z.string().email(),
  password: z.string(),
});

export const SignUpData = z.object({
  firstName: z.string(),
  lastName: z.string(),
  email: z.string().email(),
  password: z.string(),
});

export type LoginDataSchema = z.infer<typeof LoginData>;
export type SignUpDataScehema = z.infer<typeof SignUpData>;
