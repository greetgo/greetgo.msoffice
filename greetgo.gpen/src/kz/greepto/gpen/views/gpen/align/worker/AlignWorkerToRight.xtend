package kz.greepto.gpen.views.gpen.align.worker

import kz.greepto.gpen.drawport.DrawPort
import kz.greepto.gpen.drawport.Kolor
import kz.greepto.gpen.util.ColorManager

class AlignWorkerToRight implements AlignWorker {

   override paintIcon(DrawPort dp, ColorManager colors, int width, int height) {
    dp.style.foreground = Kolor.BLUE

    dp.from(60, 5).shift(0, 50).line

    dp.from(10, 10).shift(45, 20).rect.draw

    dp.from(30, 40).shift(25, 10).rect.draw

    dp.from(10, 45).shift(15, 0).line
    dp.from(20, 40).shift(5, 5).line.shift(-5,5).line
  }

}
