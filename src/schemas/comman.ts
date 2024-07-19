import { z } from "zod";

export const DefaultSearchParams = z.object({
  page: z.string().optional(),
  size: z.string().optional(),
  order: z.enum(["ASC", "DESC"]).optional(),
});
