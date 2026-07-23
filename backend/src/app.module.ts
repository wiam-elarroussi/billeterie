import { Module } from '@nestjs/common';
import { AuthController } from './auth/auth.controller';
import { AuthService } from './auth/auth.service';
import { TicketsController } from './tickets/tickets.controller';
import { TicketsService } from './tickets/tickets.service';
import { AccessControlController } from './access-control/access-control.controller';

@Module({
  imports: [],
  controllers: [AuthController, TicketsController, AccessControlController],
  providers: [AuthService, TicketsService],
})
export class AppModule {}
