import { Injectable, UnauthorizedException } from '@nestjs/common';

@Injectable()
export class AuthService {
  private users = [
    { id: 'usr-admin-01', email: 'admin@somayar.ma', name: 'Youssef (Admin Stade)', role: 'SUPER_ADMIN' },
    { id: 'usr-pos-01', email: 'pos@somayar.ma', name: 'Mehdi (Agent Caissier)', role: 'POS_AGENT' },
    { id: 'usr-pda-01', email: 'pda@somayar.ma', name: 'Samira (Agent PDA)', role: 'PDA_AGENT' },
    { id: 'usr-client-01', email: 'karim@gmail.com', name: 'Karim Bennani', role: 'CLIENT' },
  ];

  async login(email: string, pass: string) {
    const user = this.users.find(u => u.email === email);
    if (!user) {
      throw new UnauthorizedException('Identifiants invalides');
    }

    return {
      accessToken: `etk_jwt_token_${user.id}_${Date.now()}`,
      user: {
        id: user.id,
        email: user.email,
        name: user.name,
        role: user.role
      }
    };
  }

  async getProfile(userId: string) {
    return this.users.find(u => u.id === userId) || this.users[3];
  }
}
