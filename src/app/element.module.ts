import { DoBootstrap, Injector, NgModule } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { BrowserModule } from '@angular/platform-browser';
import { DAXEXLComponent } from './daxe-element/daxe-xl/daxe-xl.component';

@NgModule({
  imports: [
    BrowserModule,
    DAXEXLComponent
  ]
})
export class ElementModule implements DoBootstrap {
  constructor(private readonly injector: Injector) {}

  ngDoBootstrap(): void {
    const element = createCustomElement(DAXEXLComponent, { injector: this.injector });
    if (!customElements.get('daxe-xl')) {
      customElements.define('daxe-xl', element);
    }
  }
}
