import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def load_and_clean_data(file_path):
    """데이터 로드 및 전처리"""
    # CSV 파일 읽기
    df = pd.read_csv(file_path)
    
    print("=== 원본 데이터 정보 ===")
    print(f"총 데이터 개수: {len(df)}개")
    print(f"컬럼: {list(df.columns)}")
    print()
    
    # 데이터 정리
    # 1. 날짜 컬럼 변환
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # 2. 제품명과 카테고리 대소문자 정리
    df['ProductName'] = df['ProductName'].str.title()  # 첫 글자만 대문자
    df['Category'] = df['Category'].str.title()
    df['Salesperson'] = df['Salesperson'].str.title()
    
    # 3. 빈 값이나 잘못된 데이터 제거
    df = df.dropna(subset=['Date', 'ProductID', 'Quantity', 'UnitPrice'])
    df = df[df['ProductID'] != 'P0000']  # 잘못된 제품 ID 제거
    df = df[df['Quantity'] > 0]  # 수량이 0인 데이터 제거
    df = df[df['UnitPrice'] > 0]  # 단가가 0인 데이터 제거
    
    # 4. TotalPrice 재계산 (일부 데이터에 오류가 있을 수 있음)
    df['TotalPrice'] = df['Quantity'] * df['UnitPrice']
    
    print("=== 정리된 데이터 정보 ===")
    print(f"정리 후 데이터 개수: {len(df)}개")
    print()
    
    return df

def generate_summary_statistics(df):
    """요약 통계 생성"""
    print("="*50)
    print("📊 판매 데이터 요약 보고서")
    print("="*50)
    
    # 기본 통계
    total_sales = df['TotalPrice'].sum()
    total_quantity = df['Quantity'].sum()
    avg_order_value = df['TotalPrice'].mean()
    unique_products = df['ProductID'].nunique()
    date_range = f"{df['Date'].min().strftime('%Y-%m-%d')} ~ {df['Date'].max().strftime('%Y-%m-%d')}"
    
    print(f"📅 분석 기간: {date_range}")
    print(f"💰 총 매출액: {total_sales:,.0f}원")
    print(f"📦 총 판매 수량: {total_quantity:,}개")
    print(f"📈 평균 주문 금액: {avg_order_value:,.0f}원")
    print(f"🛍️ 판매된 제품 종류: {unique_products}개")
    print()

def analyze_by_category(df):
    """카테고리별 분석"""
    print("="*30)
    print("📊 카테고리별 판매 분석")
    print("="*30)
    
    category_sales = df.groupby('Category').agg({
        'TotalPrice': 'sum',
        'Quantity': 'sum',
        'ProductID': 'nunique'
    }).round(2)
    
    category_sales.columns = ['총 매출액', '총 수량', '제품 종류 수']
    category_sales = category_sales.sort_values('총 매출액', ascending=False)
    
    print(category_sales)
    print()
    
    return category_sales

def analyze_by_product(df):
    """제품별 분석"""
    print("="*30)
    print("🏆 베스트셀러 제품 TOP 10")
    print("="*30)
    
    product_sales = df.groupby(['ProductID', 'ProductName']).agg({
        'TotalPrice': 'sum',
        'Quantity': 'sum'
    }).round(2)
    
    product_sales.columns = ['총 매출액', '총 수량']
    top_products = product_sales.sort_values('총 매출액', ascending=False).head(10)
    
    print(top_products)
    print()
    
    return top_products

def analyze_by_region(df):
    """지역별 분석"""
    print("="*30)
    print("🌍 지역별 판매 분석")
    print("="*30)
    
    region_sales = df.groupby('Region').agg({
        'TotalPrice': 'sum',
        'Quantity': 'sum'
    }).round(2)
    
    region_sales.columns = ['총 매출액', '총 수량']
    region_sales = region_sales.sort_values('총 매출액', ascending=False)
    
    print(region_sales)
    print()
    
    return region_sales

def analyze_by_salesperson(df):
    """영업사원별 분석"""
    print("="*30)
    print("👤 영업사원별 판매 성과")
    print("="*30)
    
    # 빈 값 제거
    df_clean = df[df['Salesperson'].notna() & (df['Salesperson'] != '')]
    
    salesperson_sales = df_clean.groupby('Salesperson').agg({
        'TotalPrice': 'sum',
        'Quantity': 'sum',
        'Date': 'count'  # 거래 횟수
    }).round(2)
    
    salesperson_sales.columns = ['총 매출액', '총 수량', '거래 횟수']
    salesperson_sales = salesperson_sales.sort_values('총 매출액', ascending=False)
    
    print(salesperson_sales)
    print()
    
    return salesperson_sales

def analyze_daily_trends(df):
    """일별 판매 추이 분석"""
    print("="*30)
    print("📈 일별 판매 추이")
    print("="*30)
    
    daily_sales = df.groupby('Date').agg({
        'TotalPrice': 'sum',
        'Quantity': 'sum'
    }).round(2)
    
    daily_sales.columns = ['일별 매출액', '일별 수량']
    
    # 최고/최저 매출일
    max_sales_day = daily_sales['일별 매출액'].idxmax()
    min_sales_day = daily_sales['일별 매출액'].idxmin()
    
    print(f"최고 매출일: {max_sales_day.strftime('%Y-%m-%d')} - {daily_sales.loc[max_sales_day, '일별 매출액']:,.0f}원")
    print(f"최저 매출일: {min_sales_day.strftime('%Y-%m-%d')} - {daily_sales.loc[min_sales_day, '일별 매출액']:,.0f}원")
    print(f"일평균 매출액: {daily_sales['일별 매출액'].mean():,.0f}원")
    print()
    
    # 일별 매출 상위 5일
    print("일별 매출 TOP 5:")
    top5_days = daily_sales.sort_values('일별 매출액', ascending=False).head(5)
    for date, row in top5_days.iterrows():
        print(f"  {date.strftime('%Y-%m-%d')}: {row['일별 매출액']:,.0f}원")
    print()
    
    return daily_sales

def generate_excel_report(df, category_sales, region_sales, salesperson_sales, daily_sales, top_products):
    """Excel 보고서 생성"""
    print("="*30)
    print("📝 Excel 보고서 생성 중...")
    print("="*30)
    
    try:
        with pd.ExcelWriter('sales_analysis_report.xlsx', engine='openpyxl') as writer:
            # 원본 데이터 (정리된 버전)
            df.to_excel(writer, sheet_name='원본데이터', index=False)
            
            # 요약 통계
            summary_data = {
                '구분': ['총 매출액', '총 판매수량', '평균 주문금액', '제품 종류 수', '분석 기간'],
                '값': [
                    f"{df['TotalPrice'].sum():,.0f}원",
                    f"{df['Quantity'].sum():,}개",
                    f"{df['TotalPrice'].mean():,.0f}원",
                    f"{df['ProductID'].nunique()}개",
                    f"{df['Date'].min().strftime('%Y-%m-%d')} ~ {df['Date'].max().strftime('%Y-%m-%d')}"
                ]
            }
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='요약통계', index=False)
            
            # 각 분석 결과를 별도 시트에 저장
            category_sales.to_excel(writer, sheet_name='카테고리별분석')
            region_sales.to_excel(writer, sheet_name='지역별분석')
            salesperson_sales.to_excel(writer, sheet_name='영업사원별분석')
            daily_sales.to_excel(writer, sheet_name='일별추이')
            top_products.to_excel(writer, sheet_name='베스트셀러제품')
        
        print("✅ Excel 보고서가 'sales_analysis_report.xlsx' 파일로 저장되었습니다.")
    except Exception as e:
        print(f"❌ Excel 파일 생성 중 오류 발생: {e}")

def main():
    """메인 함수"""
    print("🚀 판매 데이터 분석을 시작합니다...\n")
    
    try:
        # 데이터 로드 및 전처리
        df = load_and_clean_data('cicd_data.csv')
        
        # 각종 분석 수행
        generate_summary_statistics(df)
        category_sales = analyze_by_category(df)
        top_products = analyze_by_product(df)
        region_sales = analyze_by_region(df)
        salesperson_sales = analyze_by_salesperson(df)
        daily_sales = analyze_daily_trends(df)
        
        # Excel 보고서 생성
        generate_excel_report(df, category_sales, region_sales, salesperson_sales, daily_sales, top_products)
        
        print("\n" + "="*50)
        print("✅ 분석이 완료되었습니다!")
        print("📊 생성된 파일:")
        print("   - sales_analysis_report.xlsx (상세 보고서)")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 분석 중 오류가 발생했습니다: {e}")

if __name__ == "__main__":
    main()