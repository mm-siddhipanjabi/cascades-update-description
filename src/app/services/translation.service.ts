import { Injectable } from '@angular/core';
import { TRANSLATIONS } from '../daxe-element/daxe-xl/daxe-header-data';

@Injectable({
  providedIn: 'root'
})
export class TranslationService {


  private lang = 'en';

  use(locale: string) {
    this.lang = locale?.split('-')[0] || 'en';
    console.log(`TranslationService: language set to ${this.lang}`);
  }

  translate(key: string, params?: Record<string, string>): string {
    let value =
      TRANSLATIONS[this.lang]?.[key] ??
      TRANSLATIONS['en']?.[key] ??
      key;

    if (params) {
      for (const p of Object.keys(params)) {
        value = value.replace(`{{${p}}}`, params[p]);
      }
    }
    return value;
  }
}
