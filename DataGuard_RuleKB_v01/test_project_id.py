"""사업 ID와 실제 검사결과서 출처가 잘못 연결되지 않는지 확인하는 테스트."""
from pathlib import Path
import unittest
from extract_rules import identify_project_id


class ProjectIdTests(unittest.TestCase):
    def test_correct_known_workbooks(self):
        self.assertEqual(identify_project_id(Path('(019)구문적정확성_검사결과서_국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트_261011.xlsx')), '019')
        self.assertEqual(identify_project_id(Path('(007)구문적정확성_검사결과서_기후 변수 임계 임산물 생산 예측 학습데이터_260904.xlsx')), '007')
        self.assertEqual(identify_project_id(Path('(013) 구문정확성_검사결과서_카자흐스탄 모국어 사용자의 한국어 음성데이터_260902.xlsx')), '013')

    def test_climate_short_8_prefix(self):
        self.assertEqual(identify_project_id(Path('(8)구문적정확성_검사결과서_기후 변수 임계 임산물 생산 예측 학습데이터_260904.xlsx')), '007')

    def test_new_project_ids(self):
        self.assertEqual(identify_project_id(Path('(099)구문정확성_검사결과서_새사업.xlsx')), '099')
        self.assertEqual(identify_project_id(Path('(100)구문정확성_검사결과서_새사업.xlsx')), '100')
        self.assertEqual(identify_project_id(Path('구문정확성_검사결과서_새사업.xlsx'), override='99'), '099')

    def test_mismatching_override_refused(self):
        with self.assertRaisesRegex(ValueError, '등록 차단'):
            identify_project_id(Path('(019)구문정확성_검사결과서_국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트.xlsx'), override='013')

    def test_conflicting_filename_refused(self):
        with self.assertRaisesRegex(ValueError, '다릅니다'):
            identify_project_id(Path('(007)구문정확성_검사결과서_카자흐스탄 모국어 사용자의 한국어 음성데이터.xlsx'))

    def test_override_matching_is_allowed(self):
        self.assertEqual(identify_project_id(Path('(019)구문정확성_검사결과서_국토안전관리원 스마트 안전장비 실증현장 이동형 CCTV 기반 AI 이벤트.xlsx'), override='19'), '019')


if __name__ == '__main__':
    unittest.main()
