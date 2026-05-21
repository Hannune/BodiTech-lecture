import win32com.client
import datetime
import os


# 시간대 시차 문제를 해결하고 정확한 로컬 날짜를 반환하는 함수
def get_local_date(received_time):
    try:
        if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
            return received_time.astimezone().date()
        return datetime.date(received_time.year, received_time.month, received_time.day)
    except Exception:
        return None


def get_and_merge_todays_ppt():
    base_dir = os.path.abspath(os.getcwd())
    temp_dir = os.path.join(base_dir, "다운로드된_개별_PPT")

    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)

    team_list = [
        ("연구기획팀", "허용민"),
        ("심혈관팀", "김태원"),
        ("급성감염팀", "한예지"),
        ("Cancer팀", "정진용"),
        ("호르몬팀", "이소희"),
        ("치료용항체팀", "김영은"),
        ("갑상선팀", "김세희"),
        ("당뇨팀", "함은선")
    ]
    submitted_teams = set()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        today = datetime.date.today()
        date_format_str = today.strftime("%y%m%d")  # 예: '260521'

        keywords = ["주간회의자료", "회의자료"]
        downloaded_files = []

        print(f"오늘 날짜({date_format_str}) 파일명 수집 및 미제출 부서 점검 시작...\n" + "=" * 50)

        old_email_count = 0

        for message in messages:
            if message.Class == 43:
                try:
                    msg_date = get_local_date(message.ReceivedTime)
                    if not msg_date:
                        continue

                    if msg_date < today:
                        old_email_count += 1
                        if old_email_count > 20:
                            break
                        continue

                    if msg_date != today:
                        continue

                    if message.Attachments.Count == 0:
                        continue

                    subject = message.Subject if message.Subject else ""
                    body = message.Body if message.Body else ""
                    sender_name = message.SenderName if message.SenderName else ""

                    for attachment in message.Attachments:
                        filename = attachment.FileName.lower()

                        if filename.endswith('.ppt') or filename.endswith('.pptx'):
                            # 1. 키워드 조건 (제목, 본문, 파일명 중 하나)
                            is_keyword_included = any(
                                kw in attachment.FileName or kw in subject or kw in body
                                for kw in keywords
                            )

                            # 2. [핵심 추가] 첨부파일명에 오늘 날짜(예: '260521')가 들어있는지 검사
                            is_date_included = date_format_str in attachment.FileName

                            # 키워드도 만족하고, 파일명에 오늘 날짜도 있어야 최종 다운로드 진행
                            if is_keyword_included and is_date_included:
                                # 제출한 부서 체크
                                for team_name, emp_name in team_list:
                                    if emp_name in sender_name:
                                        submitted_teams.add(team_name)

                                safe_filename = f"{message.ReceivedTime.strftime('%H%M%S')}_{attachment.FileName}"
                                save_path = os.path.join(temp_dir, safe_filename)

                                attachment.SaveAsFile(save_path)
                                downloaded_files.append(save_path)

                                print(f"📥 [다운로드 완료] {attachment.FileName} (보낸사람: {sender_name})")

                except Exception as e:
                    continue

        print("\n" + "=" * 50)
        print("📊 [주간 보고서 제출 현황]")

        unsubmitted_teams = [f"{team}({emp})" for team, emp in team_list if team not in submitted_teams]

        if unsubmitted_teams:
            print("🚨 미제출 부서가 있습니다:")
            for team_info in unsubmitted_teams:
                print(f"   ❌ {team_info}")
        else:
            print("✅ 모든 부서가 회의자료를 제출했습니다!")

        if not downloaded_files:
            print(f"\n오늘({date_format_str}) 날짜가 적힌 PPT 파일이 없어 병합을 생략합니다.")
            return

        print("\n" + "=" * 50)
        print(f"총 {len(downloaded_files)}개의 PPT 파일을 하나로 병합합니다...")

        merge_ppts(downloaded_files, base_dir, "주간보고병합.pptx")

    except Exception as e:
        print(f"아웃룩 작업 중 오류가 발생했습니다: {e}")


def merge_ppts(ppt_files, save_dir, output_filename):
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True

        merged_presentation = ppt_app.Presentations.Add()

        for ppt_file in ppt_files:
            print(f"🔄 병합 중... {os.path.basename(ppt_file)}")
            current_slide_count = merged_presentation.Slides.Count
            merged_presentation.Slides.InsertFromFile(ppt_file, current_slide_count)

        merged_filename = os.path.join(save_dir, output_filename)

        if os.path.exists(merged_filename):
            try:
                os.remove(merged_filename)
            except PermissionError:
                print(f"⚠️ 기존 '{output_filename}' 파일이 열려있어 다른 이름으로 저장합니다.")
                merged_filename = os.path.join(save_dir, f"주간보고병합_{datetime.datetime.now().strftime('%H%M%S')}.pptx")

        merged_presentation.SaveAs(merged_filename)
        # 문제가 있던 부분을 print문으로 정상 수정
        print(f"✅ 새 파일이 생성되었습니다: {merged_filename}")

    except Exception as e:
        print(f"PPT 병합 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    get_and_merge_todays_ppt()
