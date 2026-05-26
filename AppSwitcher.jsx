import React, { useState } from 'react';
import Step1 from './GST_Books_Step1';
import Step2 from './GST_Books_Step2';
import Step3 from './GST_Reconciliation_Step3';
import Step4 from './GSTR2A_vs_2B_Reconciliation';

export default function AppSwitcher() {
  const [app, setApp] = useState(null);
  
  if (app === 'step1') return <Step1 onBack={() => setApp(null)} />;
  if (app === 'step2') return <Step2 onBack={() => setApp(null)} />;
  if (app === 'step3') return <Step3 onBack={() => setApp(null)} />;
  if (app === 'step4') return <Step4 onBack={() => setApp(null)} />;

  return (
    <div style={{
      minHeight: '100vh', 
      backgroundColor: '#06111E', 
      display: 'flex', 
      flexDirection: 'column', 
      alignItems: 'center', 
      justifyContent: 'center', 
      color: '#D8EAF8', 
      padding: '2rem', 
      fontFamily: 'Inter, system-ui, sans-serif'
    }}>
      <h1 style={{fontSize: '2.5rem', fontWeight: '700', marginBottom: '0.5rem'}}>GST Data Extraction Suite</h1>
      <p style={{color: '#94A3B8', marginBottom: '3rem', fontSize: '1.125rem'}}>Select a tool to process your Tally Prime exports</p>
      
      <div style={{display: 'flex', gap: '2rem', flexWrap: 'wrap', justifyContent: 'center'}}>
        {/* Step 1 Card */}
        <div 
          onClick={() => setApp('step1')}
          onMouseEnter={(e) => { e.currentTarget.style.transform = 'translateY(-5px)'; e.currentTarget.style.borderColor = '#00C896'; }}
          onMouseLeave={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.borderColor = '#163050'; }}
          style={{
            background: '#0F2035', border: '1px solid #163050', borderRadius: '16px', padding: '2.5rem 2rem', width: '340px',
            cursor: 'pointer', transition: 'all 0.2s ease', boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.3)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center'
          }}
        >
          <div style={{width: '64px', height: '64px', borderRadius: '50%', background: '#00C89614', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '1.5rem'}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#00C896" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/>
            </svg>
          </div>
          <h2 style={{fontSize: '1.5rem', margin: '0 0 0.5rem 0', color: '#fff'}}>Output Tax</h2>
          <p style={{color: '#94A3B8', margin: 0, lineHeight: '1.5'}}>Extract Sales & Credit Note registers for Output Tax reconciliation.</p>
        </div>

        {/* Step 2 Card */}
        <div 
          onClick={() => setApp('step2')}
          onMouseEnter={(e) => { e.currentTarget.style.transform = 'translateY(-5px)'; e.currentTarget.style.borderColor = '#38BDF8'; }}
          onMouseLeave={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.borderColor = '#163050'; }}
          style={{
            background: '#0F2035', border: '1px solid #163050', borderRadius: '16px', padding: '2.5rem 2rem', width: '340px',
            cursor: 'pointer', transition: 'all 0.2s ease', boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.3)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center'
          }}
        >
          <div style={{width: '64px', height: '64px', borderRadius: '50%', background: '#38BDF814', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '1.5rem'}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#38BDF8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
              <polyline points="14 2 14 8 20 8"></polyline>
              <line x1="16" y1="13" x2="8" y2="13"></line>
              <line x1="16" y1="17" x2="8" y2="17"></line>
              <polyline points="10 9 9 9 8 9"></polyline>
            </svg>
          </div>
          <h2 style={{fontSize: '1.5rem', margin: '0 0 0.5rem 0', color: '#fff'}}>Input Tax Credit</h2>
          <p style={{color: '#94A3B8', margin: 0, lineHeight: '1.5'}}>Extract Purchase & Debit Note registers for ITC reconciliation.</p>
        </div>

        {/* Step 3 Card */}
        <div 
          onClick={() => setApp('step3')}
          onMouseEnter={(e) => { e.currentTarget.style.transform = 'translateY(-5px)'; e.currentTarget.style.borderColor = '#A78BFA'; }}
          onMouseLeave={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.borderColor = '#163050'; }}
          style={{
            background: '#0F2035', border: '1px solid #163050', borderRadius: '16px', padding: '2.5rem 2rem', width: '340px',
            cursor: 'pointer', transition: 'all 0.2s ease', boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.3)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center'
          }}
        >
          <div style={{width: '64px', height: '64px', borderRadius: '50%', background: '#A78BFA14', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '1.5rem'}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#A78BFA" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
              <polyline points="22 4 12 14.01 9 11.01"></polyline>
            </svg>
          </div>
          <h2 style={{fontSize: '1.5rem', margin: '0 0 0.5rem 0', color: '#fff'}}>Reconciliation</h2>
          <p style={{color: '#94A3B8', margin: 0, lineHeight: '1.5'}}>Match Tally Purchases with GSTR-2B directly and trace defaulters.</p>
        </div>

        {/* Step 4 Card */}
        <div 
          onClick={() => setApp('step4')}
          onMouseEnter={(e) => { e.currentTarget.style.transform = 'translateY(-5px)'; e.currentTarget.style.borderColor = '#FB923C'; }}
          onMouseLeave={(e) => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.borderColor = '#163050'; }}
          style={{
            background: '#0F2035', border: '1px solid #163050', borderRadius: '16px', padding: '2.5rem 2rem', width: '340px',
            cursor: 'pointer', transition: 'all 0.2s ease', boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.3)',
            display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center'
          }}
        >
          <div style={{width: '64px', height: '64px', borderRadius: '50%', background: '#FB923C14', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: '1.5rem'}}>
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#FB923C" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M4 22V4c0-.5.2-1 .6-1.4C5 2.2 5.5 2 6 2h8l6 6v14c0 .5-.2 1-.6 1.4-.4.4-.9.6-1.4.6H6c-.5 0-1-.2-1.4-.6C4.2 23 4 22.5 4 22z"/>
              <polyline points="14 2 14 8 20 8"/>
              <line x1="16" y1="13" x2="8" y2="13"/>
              <line x1="16" y1="17" x2="8" y2="17"/>
              <polyline points="10 9 9 9 8 9"/>
            </svg>
          </div>
          <h2 style={{fontSize: '1.5rem', margin: '0 0 0.5rem 0', color: '#fff'}}>GSTR-2A vs 2B</h2>
          <p style={{color: '#94A3B8', margin: 0, lineHeight: '1.5'}}>Reconcile real-time GSTR-2A with static GSTR-2B annual summaries.</p>
        </div>
      </div>
    </div>
  );
}
