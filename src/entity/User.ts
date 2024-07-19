import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  BeforeInsert,
  BeforeUpdate,
} from "typeorm";

import { encryptString } from "../common/utils";
import { CustomBaseEntity } from "./CustomBaseEntity";
@Entity()
export class User extends CustomBaseEntity {
  @PrimaryGeneratedColumn()
  id: number;

  @Column()
  firstName: string;

  @Column()
  lastName: string;

  @Column({
    unique: true,
  })
  email: string;

  @Column({
    length: 500,
    select: false,
  })
  password?: string;

  @BeforeInsert()
  @BeforeUpdate()
  async encryptPassword() {
    if (this.password) {
      this.password = await encryptString(this.password);
    }
  }
}
