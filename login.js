// Supabase 클라이언트 설정
const SUPABASE_URL = 'https://lmlpbjosdygnpqcnrwuj.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtbHBiam9zZHlnbnBxY25yd3VqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTI5MDA2NjgsImV4cCI6MjA2ODQ3NjY2OH0.Jt1Al2Sl44fSlRMAsvRw5cBuKfXcMzeYyzE774stBuQ';
const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// DOM 요소
const loginForm = document.getElementById('login-form');
const authMessage = document.getElementById('auth-message');
const showSignup = document.getElementById('show-signup');
const showResetPassword = document.getElementById('show-reset-password');
const authFormContainer = document.getElementById('auth-form');

// 로그인 처리
if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('login-email').value;
        const password = document.getElementById('login-password').value;
        const messageEl = document.getElementById('auth-message');
        const submitBtn = e.target.querySelector('button[type="submit"]');

        // [개선2] 로딩 상태 시작
        const originalText = submitBtn.textContent;
        submitBtn.disabled = true;
        submitBtn.textContent = '로그인 중...';
        messageEl.textContent = '';

        const { data, error } = await supabaseClient.auth.signInWithPassword({
            email: email,
            password: password,
        });

        if (error) {
            messageEl.textContent = '로그인 실패: ' + error.message;
            // [개선2] 로딩 상태 해제
            submitBtn.disabled = false;
            submitBtn.textContent = originalText;
        } else if (data.user) {
            // 사용자 제공 기존 로직 유지
            if (data.user.email === 'eowert72@gmail.com' || (data.user.user_metadata && data.user.user_metadata.is_admin === true)) {
                 window.location.href = 'index.html';
            } else {
                 window.location.href = 'index.html';
            }
        }
    });
}

// 회원가입 폼 보여주기
if (showSignup) {
    showSignup.addEventListener('click', (e) => {
        e.preventDefault();
        document.getElementById('auth-title').textContent = '회원가입';

        authFormContainer.innerHTML = `
            <form id="signup-form" class="space-y-5">
                <div class="space-y-2">
                    <label for="signup-email" class="text-sm font-bold text-on-surface-variant ml-1">이메일</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">mail</span>
                        <input type="email" id="signup-email" required placeholder="사용할 이메일 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <div class="space-y-2">
                    <label for="signup-password" class="text-sm font-bold text-on-surface-variant ml-1">비밀번호 (6자 이상)</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">lock</span>
                        <input type="password" id="signup-password" required minlength="6" placeholder="6자 이상 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <div class="space-y-2">
                    <label for="signup-password-confirm" class="text-sm font-bold text-on-surface-variant ml-1">비밀번호 확인</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">lock</span>
                        <input type="password" id="signup-password-confirm" required minlength="6" placeholder="비밀번호 다시 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <button type="submit" class="w-full bg-primary hover:bg-primary-hover text-white font-bold py-4 rounded-lg shadow-lg shadow-primary/25 transition-all duration-200 active:scale-[0.98] mt-4 flex items-center justify-center gap-2">
                    가입 요청
                </button>
            </form>
            <p id="auth-message" class="text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium"></p>
            <div class="mt-10 pt-8 border-t border-outline-variant/30 text-center">
                <p class="text-on-surface-variant text-sm mb-3">이미 계정이 있으신가요?</p>
                <a href="login.html" class="inline-block text-primary font-bold text-sm hover:underline hover:text-primary-hover transition-colors">로그인</a>
            </div>
        `;

        document.getElementById('signup-form').addEventListener('submit', async (ev) => {
            ev.preventDefault();
            const email = document.getElementById('signup-email').value;
            const password = document.getElementById('signup-password').value;
            const confirmPassword = document.getElementById('signup-password-confirm').value;
            const messageEl = document.getElementById('auth-message');
            const submitBtn = ev.target.querySelector('button[type="submit"]');

            // [개선3] 일치 여부 검사
            if (password !== confirmPassword) {
                messageEl.className = 'text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '비밀번호가 일치하지 않습니다.';
                return;
            }

            // [개선2] 로딩 상태 시작
            const originalText = submitBtn.textContent;
            submitBtn.disabled = true;
            submitBtn.textContent = '가입 요청 중...';

            const { data, error } = await supabaseClient.auth.signUp({
                email: email,
                password: password,
            });

            if (error) {
                messageEl.className = 'text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '회원가입 실패: ' + error.message;
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            } else {
                messageEl.className = 'text-green-600 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '회원가입 성공! 관리자 승인 후 로그인이 가능합니다.';
                ev.target.reset();
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            }
        });
    });
}

// 비밀번호 찾기 폼 보여주기
if (showResetPassword) {
    showResetPassword.addEventListener('click', (e) => {
        e.preventDefault();
        document.getElementById('auth-title').textContent = '비밀번호 찾기';

        authFormContainer.innerHTML = `
            <form id="reset-password-form" class="space-y-5">
                <div class="space-y-2">
                    <label for="reset-email" class="text-sm font-bold text-on-surface-variant ml-1">가입한 이메일</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">mail</span>
                        <input type="email" id="reset-email" required placeholder="가입 시 사용한 이메일 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <button type="submit" class="w-full bg-primary hover:bg-primary-hover text-white font-bold py-4 rounded-lg shadow-lg shadow-primary/25 transition-all duration-200 active:scale-[0.98] mt-4 flex items-center justify-center gap-2">
                    비밀번호 초기화 메일 보내기
                </button>
            </form>
            <p id="auth-message" class="text-sm text-center mt-4 min-h-[1.25rem] font-medium"></p>
            <div class="mt-10 pt-8 border-t border-outline-variant/30 text-center">
                <p class="text-on-surface-variant text-sm mb-3">비밀번호가 기억나셨나요?</p>
                <a href="login.html" class="inline-block text-primary font-bold text-sm hover:underline hover:text-primary-hover transition-colors">로그인으로 돌아가기</a>
            </div>
        `;

        document.getElementById('reset-password-form').addEventListener('submit', async (ev) => {
            ev.preventDefault();
            const email = document.getElementById('reset-email').value;
            const messageEl = document.getElementById('auth-message');
            const submitBtn = ev.target.querySelector('button[type="submit"]');
            
            // [개선2] 로딩 상태 처리
            const originalText = submitBtn.textContent;
            submitBtn.disabled = true;
            submitBtn.textContent = '메일 발송 중...';

            const { error } = await supabaseClient.auth.resetPasswordForEmail(email, {
                redirectTo: 'https://skim72.github.io/b2b-picking-admin/login.html',
            });

            if (error) {
                messageEl.className = 'text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '메일 발송 실패: ' + error.message;
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            } else {
                messageEl.className = 'text-green-600 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '비밀번호 초기화 메일이 발송되었습니다. 이메일함을 확인해주세요.';
                ev.target.reset();
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            }
        });
    });
}

// [개선1] 새 비밀번호 설정 처리 (메일 링크를 타고 왔을 때 실행됨)
supabaseClient.auth.onAuthStateChange(async (event, session) => {
    if (event === "PASSWORD_RECOVERY") {
        document.getElementById('auth-title').textContent = '새 비밀번호 설정';
        authFormContainer.innerHTML = `
            <form id="update-password-form" class="space-y-5">
                <div class="space-y-2">
                    <label for="new-password" class="text-sm font-bold text-on-surface-variant ml-1">새 비밀번호 (6자 이상)</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">lock</span>
                        <input type="password" id="new-password" required minlength="6" placeholder="새 비밀번호 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <div class="space-y-2">
                    <label for="new-password-confirm" class="text-sm font-bold text-on-surface-variant ml-1">새 비밀번호 확인</label>
                    <div class="relative">
                        <span class="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-outline text-[20px]">lock</span>
                        <input type="password" id="new-password-confirm" required minlength="6" placeholder="비밀번호 다시 입력" class="w-full pl-11 pr-4 py-3.5 bg-white border border-outline-variant rounded-lg focus:ring-2 focus:ring-primary/20 focus:border-primary transition-all duration-200 outline-none text-on-surface placeholder:text-outline/60">
                    </div>
                </div>
                <button type="submit" class="w-full bg-primary hover:bg-primary-hover text-white font-bold py-4 rounded-lg shadow-lg shadow-primary/25 transition-all duration-200 active:scale-[0.98] mt-4 flex items-center justify-center gap-2">
                    비밀번호 변경 완료
                </button>
            </form>
            <p id="auth-message" class="text-sm text-center mt-4 min-h-[1.25rem] font-medium"></p>
        `;

        document.getElementById('update-password-form').addEventListener('submit', async (ev) => {
            ev.preventDefault();
            const newPassword = document.getElementById('new-password').value;
            const confirmPassword = document.getElementById('new-password-confirm').value;
            const messageEl = document.getElementById('auth-message');
            const submitBtn = ev.target.querySelector('button[type="submit"]');

            if (newPassword !== confirmPassword) {
                messageEl.className = 'text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '비밀번호가 일치하지 않습니다.';
                return;
            }

            const originalText = submitBtn.textContent;
            submitBtn.disabled = true;
            submitBtn.textContent = '변경 중...';

            const { error } = await supabaseClient.auth.updateUser({ password: newPassword });

            if (error) {
                messageEl.className = 'text-red-500 text-sm text-center mt-4 min-h-[1.25rem] font-medium';
                messageEl.textContent = '변경 실패: ' + error.message;
                submitBtn.disabled = false;
                submitBtn.textContent = originalText;
            } else {
                alert('비밀번호가 성공적으로 변경되었습니다. 다시 로그인해주세요.');
                window.location.href = 'login.html';
            }
        });
    }
});