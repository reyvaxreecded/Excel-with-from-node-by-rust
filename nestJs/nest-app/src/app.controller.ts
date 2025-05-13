import { Controller, Get, Param } from '@nestjs/common';
import { AppService } from './app.service';

@Controller()
export class AppController {
  constructor(private readonly appService: AppService) {}

  @Get('excel/:filename')
  async getExcel(@Param('filename') filename: string): Promise<string[][]> {
    return this.appService.readExcelFile(filename);
  }
}
