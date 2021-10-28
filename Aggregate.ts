type Result = { key: string; value: number; };

module.exports = class Aggregate {
  private readonly keywords: string[] = [];
  private originList: string[][] = [];

  /* キーワードリストと回答の配列を作製する */
  constructor(sheetData:any[], key:string) {
    const list = []
    for (let sell of sheetData) {
      const dataArr:string[] = sell[key].split('、');
      this.originList.push(dataArr);
      list.push(...dataArr);
    }
    const keywords = new Set(list);
    this.keywords = Array.from(keywords);
  }

  /* 各回答数の結果を降順で返す*/
  allCounts(): Result[] {
    let result: Result = {key: '', value: 0}
    let results: Result[] = [];
    for (let keyword of this.keywords) {
      let count = 0;
      for (let v of this.originList) {
        if (v.includes(keyword)) {
          count +=1;
        }
      }
      result = {key: keyword, value: count}
      results.push(result);
    }
    return results.sort((a, b) => b.value - a.value);
  }
}
