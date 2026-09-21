/**
 * Собирает Google Форму проверки формулировок бренда из таблицы-реестра.
 *
 * Рабочая копия живёт в проекте «Голосование по позиционированию — сборка формы»
 * в Apps Script аккаунта georgekarpenko@gmail.com. Здесь — версия для репозитория.
 *
 * Как перезапустить (например, после правки текстов в реестре):
 *   script.google.com → проект → «Выполнить». Создастся НОВАЯ форма,
 *   старую надо убрать руками, иначе будет два адреса.
 */

const WORDING_SHEET_ID = "16UQuKuBiTNCLP6i9JvCQqW42J3B-2gebaa0LjwjH01M";

function createWordingForm() {
  const sh = SpreadsheetApp.openById(WORDING_SHEET_ID).getSheets()[0];
  const rows = sh.getRange(4, 1, sh.getLastRow() - 3, 5).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ""; });

  const groups = [];
  rows.forEach(function (r) {
    const item = {
      id: String(r[0]).trim(), block: String(r[1]).trim(), text: String(r[2]).trim(),
      src: String(r[3]).trim(), status: String(r[4]).trim()
    };
    if (!groups.length || groups[groups.length - 1].name !== item.block) {
      groups.push({ name: item.block, items: [] });
    }
    groups[groups.length - 1].items.push(item);
  });

  const form = FormApp.create("Формулировки бренда — проверка (Школа Ольги Карпенко)");
  form.setDescription(
    "Здесь собраны реальные тексты школы из всех источников: сайт, Notion, презентация " +
    "«Ценности школы», стратегия, плюс новые предложения. Под каждой формулировкой " +
    "написано, откуда она и работает ли сегодня.\n\n" +
    "По каждой выберите одно из трёх:\n" +
    "• Оставляем как есть — формулировка живая, ничего не трогаем.\n" +
    "• Убираем совсем — больше так про школу не говорим.\n" +
    "• Другое — меняем: впишите, на что именно.\n\n" +
    "Всего " + rows.length + " формулировок в 9 блоках, минут на двадцать. " +
    "Отвечает каждый за себя, ответы потом сводятся вместе."
  );
  form.setProgressBar(true);
  form.setCollectEmail(false);
  form.setShowLinkToRespondAgain(false);

  form.addMultipleChoiceItem()
    .setTitle("Кто отвечает")
    .setChoiceValues(["Гоша", "Ольга", "Сергей"])
    .setRequired(true);

  groups.forEach(function (g) {
    form.addPageBreakItem()
      .setTitle(g.name)
      .setHelpText("Третий вариант «Другое» — это «меняем»: впишите туда новую формулировку или что в ней править.");
    g.items.forEach(function (it) {
      form.addMultipleChoiceItem()
        .setTitle("«" + it.text + "»")
        .setHelpText(it.id + " · " + it.src + " · " + it.status)
        .setChoiceValues(["Оставляем как есть", "Убираем совсем"])
        .showOtherOption(true);
    });
  });

  form.addPageBreakItem().setTitle("Напоследок");
  form.addParagraphTextItem()
    .setTitle("Чего не хватает?")
    .setHelpText("Какие слова про школу нужны, но их нет ни в одном источнике? Пишите как есть, потом причешем.");
  form.addParagraphTextItem()
    .setTitle("Что здесь главное?")
    .setHelpText("Если бы пришлось оставить одну формулировку из всех — какую?");

  form.setDestination(FormApp.DestinationType.SPREADSHEET, WORDING_SHEET_ID);

  Logger.log("ФОРМА ДЛЯ ЗАПОЛНЕНИЯ: " + form.getPublishedUrl());
  Logger.log("РЕДАКТИРОВАТЬ ФОРМУ: " + form.getEditUrl());
  Logger.log("ВОПРОСОВ СОБРАНО: " + rows.length);
}
