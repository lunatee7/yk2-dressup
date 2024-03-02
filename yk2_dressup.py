import pandas as pd
from itertools import product
from tqdm import tqdm
from multiprocessing import Pool, cpu_count


file = "dressup.xlsx"

hostess_df = pd.read_excel(file, sheet_name="호스티스")
hair_df = pd.read_excel(file, sheet_name="머리")
dress_df = pd.read_excel(file, sheet_name="드레스")
hair_accessory_df = pd.read_excel(file, sheet_name="헤어 액세서리")
glasses_df = pd.read_excel(file, sheet_name="안경")
earring_df = pd.read_excel(file, sheet_name="귀걸이")
necklace_df = pd.read_excel(file, sheet_name="목걸이")
nail_df = pd.read_excel(file, sheet_name="네일")
ring_df = pd.read_excel(file, sheet_name="반지")
watch_df = pd.read_excel(file, sheet_name="손목시계")
bracelet_df = pd.read_excel(file, sheet_name="팔찌")

hostess_name = hostess_df['이름'].tolist()
hair_name = hair_df['이름'].tolist()
dress_name = dress_df['이름'].tolist()
hair_accessory_name = hair_accessory_df['이름'].tolist()
glasses_name = glasses_df['이름'].tolist()
earring_name = earring_df['이름'].tolist()
necklace_name = necklace_df['이름'].tolist()
nail_name = nail_df['이름'].tolist()
ring_name = ring_df['이름'].tolist()
watch_name = watch_df['이름'].tolist()
bracelet_name = bracelet_df['이름'].tolist()

hostess_sexy = dict(zip(hostess_df['이름'], hostess_df['Sexy']))
hair_sexy = dict(zip(hair_df['이름'], hair_df['Sexy']))
dress_sexy = dict(zip(dress_df['이름'], dress_df['Sexy']))
hair_accessory_sexy = dict(zip(hair_accessory_df['이름'], hair_accessory_df['Sexy']))
glasses_sexy = dict(zip(glasses_df['이름'], glasses_df['Sexy']))
earring_sexy = dict(zip(earring_df['이름'], earring_df['Sexy']))
necklace_sexy = dict(zip(necklace_df['이름'], necklace_df['Sexy']))
nail_sexy = dict(zip(nail_df['이름'], nail_df['Sexy']))
ring_sexy = dict(zip(ring_df['이름'], ring_df['Sexy']))
watch_sexy = dict(zip(watch_df['이름'], watch_df['Sexy']))
bracelet_sexy = dict(zip(bracelet_df['이름'], bracelet_df['Sexy']))

hostess_beauty = dict(zip(hostess_df['이름'], hostess_df['Beauty']))
hair_beauty = dict(zip(hair_df['이름'], hair_df['Beauty']))
dress_beauty = dict(zip(dress_df['이름'], dress_df['Beauty']))
hair_accessory_beauty = dict(zip(hair_accessory_df['이름'], hair_accessory_df['Beauty']))
glasses_beauty = dict(zip(glasses_df['이름'], glasses_df['Beauty']))
earring_beauty = dict(zip(earring_df['이름'], earring_df['Beauty']))
necklace_beauty = dict(zip(necklace_df['이름'], necklace_df['Beauty']))
nail_beauty = dict(zip(nail_df['이름'], nail_df['Beauty']))
ring_beauty = dict(zip(ring_df['이름'], ring_df['Beauty']))
watch_beauty = dict(zip(watch_df['이름'], watch_df['Beauty']))
bracelet_beauty = dict(zip(bracelet_df['이름'], bracelet_df['Beauty']))

hostess_cute = dict(zip(hostess_df['이름'], hostess_df['Cute']))
hair_cute = dict(zip(hair_df['이름'], hair_df['Cute']))
dress_cute = dict(zip(dress_df['이름'], dress_df['Cute']))
hair_accessory_cute = dict(zip(hair_accessory_df['이름'], hair_accessory_df['Cute']))
glasses_cute = dict(zip(glasses_df['이름'], glasses_df['Cute']))
earring_cute = dict(zip(earring_df['이름'], earring_df['Cute']))
necklace_cute = dict(zip(necklace_df['이름'], necklace_df['Cute']))
nail_cute = dict(zip(nail_df['이름'], nail_df['Cute']))
ring_cute = dict(zip(ring_df['이름'], ring_df['Cute']))
watch_cute = dict(zip(watch_df['이름'], watch_df['Cute']))
bracelet_cute = dict(zip(bracelet_df['이름'], bracelet_df['Cute']))

hostess_funny = dict(zip(hostess_df['이름'], hostess_df['Funny']))
hair_funny = dict(zip(hair_df['이름'], hair_df['Funny']))
dress_funny = dict(zip(dress_df['이름'], dress_df['Funny']))
hair_accessory_funny = dict(zip(hair_accessory_df['이름'], hair_accessory_df['Funny']))
glasses_funny = dict(zip(glasses_df['이름'], glasses_df['Funny']))
earring_funny = dict(zip(earring_df['이름'], earring_df['Funny']))
necklace_funny = dict(zip(necklace_df['이름'], necklace_df['Funny']))
nail_funny = dict(zip(nail_df['이름'], nail_df['Funny']))
ring_funny = dict(zip(ring_df['이름'], ring_df['Funny']))
watch_funny = dict(zip(watch_df['이름'], watch_df['Funny']))
bracelet_funny = dict(zip(bracelet_df['이름'], bracelet_df['Funny']))


def process_combination(comb):
    hostess, hair, dress, hair_accessory, glasses, earring, necklace, nail, ring, watch, bracelet = comb

    sexy = max(0, min(3, (
            hostess_sexy.get(hostess) +
            hair_sexy.get(hair) +
            dress_sexy.get(dress) +
            hair_accessory_sexy.get(hair_accessory) +
            glasses_sexy.get(glasses) +
            earring_sexy.get(earring) +
            necklace_sexy.get(necklace) +
            nail_sexy.get(nail) +
            ring_sexy.get(ring) +
            watch_sexy.get(watch) +
            bracelet_sexy.get(bracelet)
    )))

    beauty = max(0, min(3, (
            hostess_beauty.get(hostess) +
            hair_beauty.get(hair) +
            dress_beauty.get(dress) +
            hair_accessory_beauty.get(hair_accessory) +
            glasses_beauty.get(glasses) +
            earring_beauty.get(earring) +
            necklace_beauty.get(necklace) +
            nail_beauty.get(nail) +
            ring_beauty.get(ring) +
            watch_beauty.get(watch) +
            bracelet_beauty.get(bracelet)
    )))

    cute = max(0, min(3, (
            hostess_cute.get(hostess) +
            hair_cute.get(hair) +
            dress_cute.get(dress) +
            hair_accessory_cute.get(hair_accessory) +
            glasses_cute.get(glasses) +
            earring_cute.get(earring) +
            necklace_cute.get(necklace) +
            nail_cute.get(nail) +
            ring_cute.get(ring) +
            watch_cute.get(watch) +
            bracelet_cute.get(bracelet)
    )))

    funny = max(0, min(3, (
            hostess_funny.get(hostess) +
            hair_funny.get(hair) +
            dress_funny.get(dress) +
            hair_accessory_funny.get(hair_accessory) +
            glasses_funny.get(glasses) +
            earring_funny.get(earring) +
            necklace_funny.get(necklace) +
            nail_funny.get(nail) +
            ring_funny.get(ring) +
            watch_funny.get(watch) +
            bracelet_funny.get(bracelet)
    )))

    count_3 = [sexy, beauty, cute, funny].count(3)
    count_0 = [sexy, beauty, cute, funny].count(0)

    if count_3 == 3 and count_0 == 1:
        return {
            '호스티스': hostess,
            '머리': hair,
            '드레스': dress,
            '헤어 액세서리': hair_accessory,
            '안경': glasses,
            '귀걸이': earring,
            '목걸이': necklace,
            '네일': nail,
            '반지': ring,
            '손목시계': watch,
            '팔찌': bracelet,
            'Sexy': sexy,
            'Beauty': beauty,
            'Cute': cute,
            'Funny': funny
        }
    else:
        return None


if __name__ == "__main__":
    num_processes = cpu_count()
    with Pool(num_processes) as pool:
        comb_result = list(tqdm(pool.imap(process_combination, product(
            hostess_name,
            hair_name,
            dress_name,
            hair_accessory_name,
            glasses_name,
            earring_name,
            necklace_name,
            nail_name,
            ring_name,
            watch_name,
            bracelet_name
        )), total=(
                len(hostess_name) *
                len(hair_name) *
                len(dress_name) *
                len(hair_accessory_name) *
                len(glasses_name) *
                len(earring_name) *
                len(necklace_name) *
                len(nail_name) *
                len(ring_name) *
                len(watch_name) *
                len(bracelet_name)
        )))

    comb_result = [result for result in comb_result if result is not None]

    result_df = pd.DataFrame(comb_result)

    output_path = "YK2_Dressup.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        result_df[result_df['호스티스'] == '코유키'].to_excel(writer, index=False, sheet_name='코유키')
        result_df[result_df['호스티스'] == '카나'].to_excel(writer, index=False, sheet_name='카나')
        result_df[result_df['호스티스'] == 'AIKA'].to_excel(writer, index=False, sheet_name='AIKA')
        result_df[result_df['호스티스'] == '쇼코'].to_excel(writer, index=False, sheet_name='쇼코')
        result_df[result_df['호스티스'] == '유아'].to_excel(writer, index=False, sheet_name='유아')
        result_df[result_df['호스티스'] == '키라라'].to_excel(writer, index=False, sheet_name='키라라')
