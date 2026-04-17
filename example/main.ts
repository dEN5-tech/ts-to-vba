/**
 * Базовый класс транзакции
 */
export class BaseTransaction {
    constructor(
        public id: string,
        public description: string,
        public amount: number
    ) {}

    // Метод, который будет переопределен
    public getFormattedType(): string {
        return "GENERIC";
    }

    public logInfo(): void {
        console.log(`Транзакция ${this.id}: ${this.description}`);
    }
}

/**
 * Класс ДОХОД (Наследует BaseTransaction)
 */
export class Income extends BaseTransaction {
    constructor(id: string, description: string, amount: number, public source: string) {
        super(id, description, amount);
    }

    // Переопределение метода (Override)
    public getFormattedType(): string {
        return `ДОХОД (${this.source})`;
    }
}

/**
 * Класс РАСХОД (Наследует BaseTransaction)
 */
export class Expense extends BaseTransaction {
    constructor(id: string, description: string, amount: number, public category: string) {
        super(id, description, amount);
    }

    // Переопределение метода (Override)
    public getFormattedType(): string {
        return `РАСХОД [${this.category}]`;
    }

    // Уникальный метод расхода
    public getTaxDeduction(): number {
        return this.amount * 0.13; // 13% вычет
    }
}

/**
 * Главная функция теста
 */
export function runInheritanceReport(): void {
    const sheet: any = (globalThis as any).ActiveSheet;
    const transactions: BaseTransaction[] = [];

    try {
        console.log("Запуск теста наследования...");

        // Создаем разные типы объектов в один массив (Полиморфизм!)
        transactions.push(new Income("INC01", "Зарплата", 5000, "Работа"));
        transactions.push(new Expense("EXP01", "Аренда", 1500, "Жилье"));
        transactions.push(new Expense("EXP02", "Продукты", 200, "Еда"));

        sheet.Range("A1:D1").Value = ["ID", "Тип", "Описание", "Сумма"];
        sheet.Range("A1:D1").Font.Bold = true;

        let row = 2;
        for (const tx of transactions) {
            sheet.Cells(row, 1).Value = tx.id;
            sheet.Cells(row, 2).Value = tx.getFormattedType(); // Вызов переопределенного метода
            sheet.Cells(row, 3).Value = tx.description;
            sheet.Cells(row, 4).Value = tx.amount;

            // Проверка специфичного метода (instanceof-like в VBA)
            if (tx.getFormattedType().includes("РАСХОД")) {
                sheet.Cells(row, 4).Font.Color = 255; // Красный для расходов
            } else {
                sheet.Cells(row, 4).Font.Color = 32768; // Зеленый для доходов
            }

            row++;
        }

        console.log("Тест наследования завершен успешно.");

    } catch (e) {
        console.log("Ошибка в тесте наследования:", e);
    }
}