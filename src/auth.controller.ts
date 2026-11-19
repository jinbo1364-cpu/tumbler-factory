import { Body, Controller, Get, Post, Req } from '@nestjs/common';
import type { Request } from 'express';
import 'express-session';

const users = {
  writer: { password: 'writer1234!', role: 'writer' },
  reviewer: { password: 'reviewer1234!', role: 'reviewer' },
  'factory-user': { password: 'factory1234!', role: 'factory-user' },
} as const;

type UserName = keyof typeof users;

declare module 'express-session' {
  interface SessionData {
    user?: { username: UserName; role: string };
  }
}

@Controller('api')
export class AuthController {
  @Get('session')
  getSession(@Req() request: Request) {
    const user = request.session.user;
    return user ? { user } : { user: null };
  }

  @Post('login')
  login(
    @Body() credentials: { username?: string; password?: string },
    @Req() request: Request,
  ) {
    const { username, password } = credentials;
    if (
      !username ||
      !(username in users) ||
      users[username as UserName].password !== password
    ) {
      return { error: '로그인 정보를 확인하세요.' };
    }

    const user = {
      username: username as UserName,
      role: users[username as UserName].role,
    };
    request.session.user = user;
    return { user };
  }

  @Post('logout')
  logout(@Req() request: Request) {
    delete request.session.user;
    return { user: null };
  }
}