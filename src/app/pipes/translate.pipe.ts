import { Pipe, PipeTransform } from '@angular/core';
import { TranslationService } from '../services/translation.service';

@Pipe({
  name: 'translate',
  standalone: true,
  pure: false // IMPORTANT
})
export class TranslatePipe implements PipeTransform {
  constructor(private i18n: TranslationService) { }
  transform(key: string, params?: Record<string, string>): string {
    return this.i18n.translate(key, params);
  }

}
