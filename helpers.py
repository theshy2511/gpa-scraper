import re

def extract_gpa(soup):
    """
    HTML thực tế:
    <td class="text-start">Trung bình chung tích lũy:</td>
    <td><span>6,82</span> - <span>2,61</span> 30</td>

    => lấy span thứ 2 (GPA hệ 4)
    """
    try:
        # tìm td chứa label
        label_td = soup.find(
            "td",
            string=re.compile(r"trung\s*bình\s*chung\s*tích\s*lu", re.IGNORECASE)
        )
        if not label_td:
            return None

        # td kế bên chứa số
        value_td = label_td.find_next_sibling("td")
        if not value_td:
            return None

        spans = value_td.find_all("span")
        if len(spans) < 2:
            return None

        # span[1] = GPA hệ 4
        gpa_str = spans[1].get_text(strip=True).replace(",", ".")
        gpa = float(gpa_str)

        # validate hệ 4
        if 0 <= gpa <= 4:
            return gpa

        return None

    except Exception:
        return None
